"""
This module provides functionality to parse SQL files and extract queries along with their associated configuration options for Excel export.
"""

# TODO reading a SQL file assumes all queries are ready for execution (no support for parameterization)


import os
import re
from itertools import product
from typing import List

from sql2excel.sqlexec import QueryConfig

_data_columns_pattern = r"data_columns\s*[:=]\s*[\[\(]\s*[^)\]]*[\)\]]"
_headings_pattern = r"headings\s*[:=]\s*[\[\(]\s*[^)\]]*[\)\]]"


def _convert(value):
    """Convert a string to the most specific number type or return it unchanged.

    Parameters
    ----------
    value : object
        The value to convert. Non-string values are returned as-is.

    Returns
    -------
    bool, int, float, or str
        The converted value. Tries ``bool``, then ``int``, then ``float``;
        falls back to the original string.
    """
    if not isinstance(value, str):
        return value

    if value.lower() == "false":
        return False

    if value.lower() == "true":
        return True

    # Convert to the most specific numeric type
    try:
        return int(value)
    except ValueError:
        pass

    try:
        return float(value)
    except ValueError:
        pass

    return value


def _from_pipe_sep_values_to_list_of_dict(text):
    """Convert a pipe-separated key-value string to a list of single-key dicts.

    Parameters
    ----------
    text : str
        A string of the form ``"key = value1 | value2 | value3"``.
        The key and values can be separated by ``=`` or ``:`` with optional
        surrounding whitespace.

    Returns
    -------
    list of dict or object
        A list such as ``[{key: value1}, {key: value2}, ...]`` where each
        value is passed through ``_convert``. Returns *text* unchanged if it
        does not match the expected format or is not a string.
    """
    if not isinstance(text, str):
        return text

    # Split the string by the first occurrence of '=' or ':'
    parts = re.split(r"\s*[:=]\s*", text, 1)
    if len(parts) < 2:
        return text  # Not in the expected format

    key = parts[0].strip()
    values_str = parts[1].strip()

    # Split the values string by '|'
    values = [v.strip() for v in values_str.split("|")]

    result = []
    for val in values:
        result.append({key: _convert(val)})

    return result


def _parse_list_or_tuple(text):
    """Parse a string representation of a Python list or tuple.

    Parameters
    ----------
    text : str
        A string like ``"[1, 2, 3]"`` or ``(1, 2, 3)``. Elements are
        automatically converted via ``_convert``.

    Returns
    -------
    list, tuple, or object
        A ``list`` if *text* starts with ``[``, a ``tuple`` if it starts
        with ``(``, otherwise *text* is returned unchanged.
    """
    pattern = r"^[\[\(]\s*(.*)\s*[\]\)]$"
    match = re.match(pattern, text)

    if match:
        elements = match.group(1)

        elements = [
            _convert(elem.strip().strip("'\""))
            for elem in re.split(r"\s*,\s*", elements)
            if elem
        ]
        if text.startswith("["):
            return elements
        elif text.startswith("("):
            return tuple(elements)

    return text


def parse_sql_file(filepath: str) -> List[QueryConfig]:
    """Parse an annotated SQL file into a list of ``QueryConfig`` objects.

    The file is split on semicolons into individual queries.  Comment lines
    (starting with ``--``) are interpreted as directives that configure
    chart type, title, section heading, and other Excel parameters.
    Supported directives include ``chart``, ``title``, ``sheet_param``,
    ``query_param``, and any other ``key=value`` pair consumed by
    ``QueryConfig``.

    Parameters
    ----------
    filepath : str
        Path to the ``.sql`` file to parse.

    Returns
    -------
    list of QueryConfig
        One ``QueryConfig`` per query (or per combination of
        ``query_param`` / ``sheet_param`` values).

    Raises
    ------
    ValueError
        If ``sheet_param`` directives contain duplicate parameter names,
        if more than one ``sheet_param`` name is used, or if option
        formatting is invalid.
    """
    with open(filepath, "r") as file:
        content = file.read()

    # Split the content by semicolons to separate queries
    raw_queries = content.split(";")

    queries_config = []
    # Each entry is a dict: {param_name: [val1, val2, ...]}
    sheet_params = []
    for query in raw_queries:

        if not query.strip():
            continue

        # Split each segment by lines
        lines = query.splitlines()

        # Extract comments for QueryConfig
        excel_options = {}
        # A list of list of dictionaries
        query_params = []
        for line in lines:
            if line.strip().startswith("--"):
                line = line.replace("--", "")

                # Sheet parameters
                if "sheet_param" in line:
                    line = line.replace("sheet_param", "")
                    parsed = _from_pipe_sep_values_to_list_of_dict(line)
                    if isinstance(parsed, list) and len(parsed) > 0:
                        param_dict = parsed[0]
                        param_name = list(param_dict.keys())[0]
                        param_values = [list(d.values())[0] for d in parsed]
                        # Check for duplicate parameter names
                        existing_names = [list(sp.keys())[0] for sp in sheet_params]
                        if param_name in existing_names:
                            raise ValueError(
                                f"Duplicate sheet_param name '{param_name}' is not allowed"
                            )
                        sheet_params.append({param_name: param_values})
                    continue

                # Query parameters
                if "query_param" in line:
                    line = line.replace("query_param", "")
                    qp = _from_pipe_sep_values_to_list_of_dict(line)
                    query_params.append(qp)
                    continue

                # Arguments that are list-like or tuple-like
                data_columns = re.findall(_data_columns_pattern, line)
                data_columns = data_columns[0] if data_columns else []
                line = re.sub(_data_columns_pattern, "", line)
                headings = re.findall(_headings_pattern, line)
                headings = headings[0] if headings else []
                line = re.sub(_headings_pattern, "", line)

                # Parse the comment
                options = line.split(",")
                # options += data_columns
                if data_columns:
                    options.append(data_columns)

                if headings:
                    options.append(headings)
                try:
                    for option in options:
                        option = (
                            option.split(":") if ":" in option else option.split("=")
                        )

                        # Skips irrelevant comments
                        # `'chart'` is an exception because it can have no value
                        if (
                            len(option) < 2
                            and option[0].strip() != "chart"
                            and option[0].strip() != "exec"
                        ):
                            # print("skipped: ", option)
                            continue

                        key = option[0].strip().lower().replace(" ", "_")
                        try:
                            value = option[1].strip()
                        except IndexError:
                            value = "chart"

                        if value.startswith("[") or value.startswith("("):
                            value = _parse_list_or_tuple(value)

                        value = _convert(value)

                        if key:
                            excel_options[key] = value
                except Exception as e:
                    raise ValueError(
                        f"Incorrectly formatted SQL file. Error parsing options: {e}"
                    )

        # Create QueryConfig object
        query = query.strip() + ";"
        if query_params:
            query_params = [
                {k: v for d in combination for k, v in d.items()}
                for combination in product(*query_params)
            ]

            for qp in query_params:
                query_config = QueryConfig(
                    sql=query, from_sql_script=True, **excel_options
                )
                query_config.sql_params = qp

                # Replace parameter name in section_heading for query_param
                if "section_heading" in query_config.xl_params:
                    heading = str(query_config.xl_params["section_heading"])
                    for qp_name, qp_value in qp.items():
                        heading = heading.replace(":" + qp_name, str(qp_value))
                        heading = heading.replace(qp_name, str(qp_value))
                    query_config.xl_params["section_heading"] = heading

                queries_config.append(query_config)
        else:
            query_config = QueryConfig(sql=query, from_sql_script=True, **excel_options)

            queries_config.append(query_config)

    # Create multiple copies of query_config for each combination of sheet parameters
    if sheet_params:
        # Build the cartesian product of all parameter value lists
        param_names = [list(sp.keys())[0] for sp in sheet_params]
        param_value_lists = [list(sp.values())[0] for sp in sheet_params]

        # NOTE: sheet parameters should be used to repeat the same analysis for different entities (e.g. countries, etc)
        if len(param_names) > 1:
            raise ValueError(
                "Multiple `sheet_param` names are not supported. Please use only one `sheet_param`."
            )

        queries_config_extended = []
        for combination in product(*param_value_lists):
            param_combo = dict(zip(param_names, combination))

            for query_config in queries_config:
                # Avoid modifying the original objects which is reused across iterations
                new_xl_params = query_config.xl_params.copy()

                new_sql_params = (
                    query_config.sql_params.copy() if query_config.sql_params else {}
                )

                new_config = QueryConfig(
                    sql=query_config.sql,
                    sql_params=new_sql_params,
                    from_sql_script=query_config.from_sql_script,
                    **new_xl_params,
                )

                # Replace :param_name in SQL and add to sql_params
                for sp_name, sp_value in param_combo.items():
                    new_config.sql_params[sp_name] = sp_value

                # Use first combination's value as the sheet name
                if "sheetname" not in new_config.xl_params:
                    new_config.xl_params["sheetname"] = str(
                        list(param_combo.values())[0]
                    )

                # Replace parameter name in section_heading (with or without colon)
                if "section_heading" in new_config.xl_params:
                    heading = str(new_config.xl_params["section_heading"])
                    for sp_name, sp_value in param_combo.items():
                        heading = heading.replace(":" + sp_name, str(sp_value))
                        heading = heading.replace(sp_name, str(sp_value))
                    new_config.xl_params["section_heading"] = heading

                queries_config_extended.append(new_config)
        queries_config = queries_config_extended

    return queries_config
