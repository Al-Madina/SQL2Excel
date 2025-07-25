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
    """
    Convert all '?' placeholders in the SQL query to named parameters like ':param_1', ':param_2'
    """
    placeholder = re.compile(r"\?")

    for i in range(len(params)):
        query = placeholder.sub(f":param_{i+1}", query, count=1)

    return query


def _bind_positional_parameters(sql, params):
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

        Parameters:
        -----------
        sql : str, optional
            The SQL query to execute.
        sql_params : dict | Sequence, optional
            Parameters to be passed to the SQL query.
        from_sql_script:
            Whether this query config is read from a SQL script
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
    def __init__(
        self, conn=None, session=None, engine=None, connection_string=None, silent=False
    ):
        """
        Initialize the SQLExecutor with a database connection.
        """
        self.conn = conn
        self.session = session
        self.engine = engine
        self.silent = silent
        self.closed = None

        if not any([conn, session, engine, connection_string]):
            raise ValueError("Cannot establish a connection to the database")

        if connection_string:
            self.engine = create_engine(connection_string)
            self.conn = self.engine.connect()

        if engine:
            self.conn = self.engine.connect()

        self.closed = False

    def execute(self, query_config):
        """
        Execute the SQL query defined in the QueryConfig object and return a DataFrame.

        Parameters
        ----------
        query_config: QueryConfig
            An instance of QueryConfig containing the SQL query and parameters.
        """
        # Validate query
        if query_config.sql is None:
            raise ValueError("No SQL query is found in QueryConfig")

        sql = query_config.sql
        params = query_config.sql_params
        result = None

        try:
            """NOTE sqlalchemy 2.0 does not support positional parameters
            https://github.com/sqlalchemy/sqlalchemy/issues/5178
            Positional parameters are easier to work with even though named
            parameters are more verbose. The issue above suggests using `exec_driver_sql`.
            However, passing parameters this way is not DB-agnostic. An alternative
            way of handling this is to use bind.

            See Also
            --------
            _bind_positional_parameters
            """

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

        if result:
            columns = result.keys()
            data = result.fetchall()
            df = pd.DataFrame(data, columns=columns)
            return df

    def executeall(self, query_configs):
        results = []
        for query in query_configs:
            results.append(self.execute(query))
        return results

    def close(self):
        if self.conn:
            self.conn.close()
        if self.session:
            self.session.close()
        if self.engine:
            self.engine.dispose()
        self.closed = True
