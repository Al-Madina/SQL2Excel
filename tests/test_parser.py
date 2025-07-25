import os
import tempfile

import pytest

from sql2excel.parser import _convert, _parse_list_or_tuple, parse_sql_file
from sql2excel.sqlexec import QueryConfig


@pytest.mark.parametrize(
    "value, expected",
    [
        ("false", False),
        ("true", True),
        ("42", 42),
        ("3.14", 3.14),
        ("random_string", "random_string"),
        (None, None),
        (5, 5),
    ],
)
def test_convert(value, expected):
    assert _convert(value) == expected


@pytest.mark.parametrize(
    "text, expected",
    [
        ("[1, 2, 3]", [1, 2, 3]),
        ("(4, 5, 6)", (4, 5, 6)),
        ("['a', 'b', 'c']", ["a", "b", "c"]),
        ("()", ()),
        ("[]", []),
        ("not_a_list", "not_a_list"),
    ],
)
def test_parse_list_or_tuple(text, expected):
    assert _parse_list_or_tuple(text) == expected


def test_parse_sql_file_basic_query():
    sql_content = """
    -- chart:bar, data_column_start=2, data_column_end=4
    -- title=Rental Rate Statistics, ylabel=Rental Rates
    SELECT name, MIN(rental_rate), AVG(rental_rate), MAX(rental_rate)
    FROM category
    GROUP BY name;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()

        configs = parse_sql_file(temp_file.name)

        assert len(configs) == 1
        query_config = configs[0]
        assert isinstance(configs[0], QueryConfig)
        assert query_config.xl_params.get("chart") == "bar"
        assert query_config.xl_params.get("data_column_start") == 2
        assert query_config.xl_params.get("data_column_end") == 4
        assert query_config.xl_params.get("title") == "Rental Rate Statistics"
        assert query_config.xl_params.get("ylabel") == "Rental Rates"

        # assert the query is correctly parsed
        query_config.sql == """SELECT name, MIN(rental_rate), AVG(rental_rate), MAX(rental_rate)
        FROM category
        GROUP BY name;"""

    os.remove(temp_file.name)


def test_parse_sql_file_multiple_queries():
    sql_content = """
    -- chart:bar
    -- title=Min-Max rental rate, vary_color=True
    SELECT name, MIN(rental_rate), MAX(rental_rate)
    FROM category
    GROUP BY name;

    -- chart:line, xlabel=Date, ylabel=Rental Count
    SELECT rental_date :: date AS "Rental Date", COUNT(*) AS "Rental Count"
    FROM rental
    GROUP BY 1
    ORDER BY 1;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()

        configs = parse_sql_file(temp_file.name)

        assert len(configs) == 2

        # First query
        assert configs[0].xl_params.get("chart") == "bar"
        assert configs[0].xl_params.get("title") == "Min-Max rental rate"
        assert configs[0].xl_params.get("vary_color") is True
        configs[
            0
        ].sql == """SELECT name, MIN(rental_rate), MAX(rental_rate)
        FROM category
        GROUP BY name;"""

        # Second query
        assert configs[1].xl_params.get("chart") == "line"
        assert configs[1].xl_params.get("xlabel") == "Date"
        assert configs[1].xl_params.get("ylabel") == "Rental Count"
        configs[
            1
        ].sql == """SELECT rental_date :: date AS "Rental Date", COUNT(*) AS "Rental Count"
        FROM rental
        GROUP BY 1
        ORDER BY 1;"""

    os.remove(temp_file.name)
