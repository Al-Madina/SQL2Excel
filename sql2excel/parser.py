"""
This module provides functionality to parse SQL files and extract queries along with their associated configuration options for Excel export.
"""

# TODO reading a SQL file assumes all queries are ready for execution (no support for parameterization)


import os
import re
from typing import List

from sql2excel.sqlexec import QueryConfig

_data_columns_pattern = r"data_columns\s*[:=]\s*[\[\(]\s*[^)\]]*[\)\]]"
_headings_pattern = r"headings\s*[:=]\s*[\[\(]\s*[^)\]]*[\)\]]"


def _convert(value):
    """
    Convert a string to the most specific number class or return it as is.
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


def _parse_list_or_tuple(text):
    """
    Parse a string representing a list or tuple into the corresponding Python object.
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
    """
    Parse SQL file to extract queries and their associated QueryConfig settings.
    """
    with open(filepath, "r") as file:
        content = file.read()

    # Split the content by semicolons to separate queries
    raw_queries = content.split(";")

    queries_config = []
    for query in raw_queries:

        if not query.strip():
            continue

        # Split each segment by lines
        lines = query.split(os.linesep)

        # Extract comments for QueryConfig
        excel_options = {}
        for line in lines:
            if line.strip().startswith("--"):
                line = line.replace("--", "")
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
                        if len(option) < 2 and option[0].strip() != "chart":
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
                except:
                    raise ValueError("Incorrectly formated SQL file")

        # Create QueryConfig object
        query = query.strip() + ";"
        query_config = QueryConfig(sql=query, from_sql_script=True, **excel_options)

        queries_config.append(query_config)

    return queries_config
