"""
Module for executing SQL queries and returning results as pandas DataFrames.

Provides utilities for handling positional and named parameters in SQL queries,
and configuration for exporting query results to Excel charts.
"""

import re
import traceback
import warnings
from typing import Sequence

import pandas as pd
from sqlalchemy import bindparam, create_engine
from sqlalchemy.sql import text


def _convert_query_placeholders(query: str, params: Sequence) -> str:
    """Convert positional ``?`` placeholders to named ``:param_N`` placeholders.

    Parameters
    ----------
    query : str
        The SQL query string containing ``?`` placeholders.
    params : Sequence
        The parameter values; its length determines how many placeholders
        are replaced.

    Returns
    -------
    str
        The query with ``?`` replaced by ``:param_1``, ``:param_2``, etc.
    """
    placeholder = re.compile(r"\?")

    for i in range(len(params)):
        query = placeholder.sub(f":param_{i+1}", query, count=1)

    return query


def _bind_positional_parameters(sql, params):
    """Bind positional parameters to a SQL query using SQLAlchemy ``bindparam``.

    Parameters
    ----------
    sql : str
        The SQL query string containing ``?`` placeholders.
    params : Sequence
        The positional parameter values.

    Returns
    -------
    tuple of (sqlalchemy.sql.elements.TextClause, dict)
        A 2-tuple of *(compiled_sql, params_dict)* ready for execution.

    See Also
    --------
    _convert_query_placeholders : Converts ``?`` to named placeholders.
    """

    sql = _convert_query_placeholders(sql, params)
    params = {f"param_{i+1}": value for i, value in enumerate(params)}
    sql = text(sql)
    sql = sql.bindparams(
        *[
            (
                bindparam(key=k, value=v, expanding=True)
                if isinstance(v, Sequence) and not isinstance(v, str)
                else bindparam(key=k, value=v)
            )
            for k, v in params.items()
        ]
    )
    return sql, params


class QueryConfig:

    def __init__(
        self, sql=None, sql_params=None, from_sql_script=False, **xl_params
    ) -> None:
        """
        Initialize the object with SQL query, parameters, and additional Excel parameters.

        Parameters
        ----------
        sql : str, optional
            The SQL query to execute.
        sql_params : dict or Sequence, optional
            Parameters to be passed to the SQL query.
        from_sql_script : bool, optional
            Whether this query config was read from a SQL script.
            Default is ``False``.
        **xl_params : dict
            Additional parameters for configuring the Excel chart. The keys can include:

            - chart : str
                The chart class. E.g. line for LineChart, pie for PieChart, etc.
            - section_heading : str
                The heading placed directly above the data and chart.
            - title : str
                The title of the chart.
            - ylabel : str
                The label for the y-axis.
            - xlabel : str
                The label for the x-axis.
            - rotation : int
                The rotation angle for the x-axis labels.
            - xlim : tuple
                The limits for the x-axis (min, max).
            - ylim : tuple
                The limits for the y-axis (min, max).
            - y_orientation : str
                The orientation of the y-axis ('minMax', 'maxMin').
            - width : int
                The width of the chart.
            - height : int
                The height of the chart.
            - no_legend : bool
                Whether to hide the legend.
            - legend_position : str
                The position of the legend.
            - chart_position : str
                The position of the chart ('right', 'bottom').
        """
        self.sql = sql
        self.sql_params = sql_params
        self.from_sql_script = from_sql_script
        self.xl_params = xl_params


class SQLExecutor:
    def __init__(self, conn=None, session=None, engine=None, db_url=None, silent=False):
        """Initialize the SQLExecutor with a database connection.

        Exactly one of *conn*, *session*, *engine*, or *db_url* must be
        provided.

        Parameters
        ----------
        conn : sqlalchemy.engine.Connection, optional
            An existing database connection.
        session : sqlalchemy.orm.Session, optional
            An existing SQLAlchemy session.
        engine : sqlalchemy.engine.Engine, optional
            An existing SQLAlchemy engine; a connection will be created from it.
        db_url : str, optional
            A SQLAlchemy database URL used to create a new engine and connection.
        silent : bool, optional
            If ``True``, query execution errors are caught and only produce
            warnings instead of raising.

        Raises
        ------
        ValueError
            If none of *conn*, *session*, *engine*, or *db_url* is provided.
        """
        self.conn = conn
        self.session = session
        self.engine = engine
        self.silent = silent
        self.closed = None

        if not any([conn, session, engine, db_url]):
            raise ValueError("Cannot establish a connection to the database")

        if db_url:
            self.engine = create_engine(db_url)
            self.conn = self.engine.connect()
        elif engine:
            self.conn = self.engine.connect()
        elif conn:
            pass
        elif session:
            pass

        self.closed = False

    def execute(self, query_config):
        """Execute the SQL query defined in the QueryConfig object and return a DataFrame.

        Parameters
        ----------
        query_config : QueryConfig
            An instance of QueryConfig containing the SQL query and parameters.

        Returns
        -------
        pandas.DataFrame or None
            A DataFrame with the query results, or ``None`` if the statement
            does not produce rows.

        Raises
        ------
        ValueError
            If ``query_config.sql`` is ``None``.
        Exception
            Re-raises any database exception when *silent* is ``False``.
        """
        # Validate query
        if query_config.sql is None:
            raise ValueError("No SQL query is found in QueryConfig")

        sql = query_config.sql
        params = query_config.sql_params
        result = None

        try:
            # NOTE sqlalchemy 2.0 does not support positional parameters
            # https://github.com/sqlalchemy/sqlalchemy/issues/5178
            # Positional parameters are easier to work with even though named
            # parameters are more verbose. The issue above suggests using
            # `exec_driver_sql`.  However, passing parameters this way is not
            # DB-agnostic. An alternative way of handling this is to use bind.
            #
            # See Also
            # --------
            # _bind_positional_parameters

            if isinstance(params, Sequence):
                sql, params = _bind_positional_parameters(sql, params)

            else:
                sql = text(sql)

            if self.conn:
                result = self.conn.execute(sql, params)
            elif self.session:
                result = self.session.execute(sql, params)
            else:
                pass

        except Exception as e:
            if self.silent:
                warnings.warn(
                    f"Unable to execute the query due to the exception below: {sql}"
                )
                traceback.print_exc()
            else:
                raise e

        try:
            columns = result.keys()
            data = result.fetchall()
            df = pd.DataFrame(data, columns=columns)
            return df
        # Raises sqlalchemy.exc.ResourceClosedError if sql stmt does not produce rows
        except Exception:
            return None

    def executeall(self, query_configs):
        """Execute a sequence of QueryConfig objects and collect the results.

        Parameters
        ----------
        query_configs : Sequence[QueryConfig]
            A sequence of QueryConfig objects to execute.

        Returns
        -------
        list of pandas.DataFrame or None
            One result per query, in the same order.
        """
        results = []
        for query in query_configs:
            results.append(self.execute(query))
        return results

    def close(self):
        """Close all open database resources.

        Disposes of the connection, session, and engine (whichever are set)
        and marks the executor as closed.
        """
        if self.conn:
            self.conn.close()
        if self.session:
            self.session.close()
        if self.engine:
            self.engine.dispose()
        self.closed = True
