import os
import tempfile

import pytest

from sql2excel.parser import (
    _convert,
    _from_pipe_sep_values_to_list_of_dict,
    _parse_list_or_tuple,
    parse_sql_file,
)
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


@pytest.mark.parametrize(
    "text, expected",
    [
        ("key = value1 | value2", [{"key": "value1"}, {"key": "value2"}]),
        ("key=value1|value2", [{"key": "value1"}, {"key": "value2"}]),
        ("key : value1 | value2", [{"key": "value1"}, {"key": "value2"}]),
        ("key:value1|value2", [{"key": "value1"}, {"key": "value2"}]),
        (
            "another_key: 1 | 2 | 3",
            [{"another_key": 1}, {"another_key": 2}, {"another_key": 3}],
        ),
        ("single_value = test", [{"single_value": "test"}]),
        ("no_separator", "no_separator"),
        (123, 123),
        ("key = true | false", [{"key": True}, {"key": False}]),
        ("key = 1.1 | 2.2", [{"key": 1.1}, {"key": 2.2}]),
    ],
)
def test_from_pipe_sep_values_to_list_of_dict(text, expected):
    assert _from_pipe_sep_values_to_list_of_dict(text) == expected


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
        assert (
            query_config.sql == """-- chart:bar, data_column_start=2, data_column_end=4
    -- title=Rental Rate Statistics, ylabel=Rental Rates
    SELECT name, MIN(rental_rate), AVG(rental_rate), MAX(rental_rate)
    FROM category
    GROUP BY name;"""
        )

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
        assert configs[0].sql == """-- chart:bar
    -- title=Min-Max rental rate, vary_color=True
    SELECT name, MIN(rental_rate), MAX(rental_rate)
    FROM category
    GROUP BY name;"""

        # Second query
        assert configs[1].xl_params.get("chart") == "line"
        assert configs[1].xl_params.get("xlabel") == "Date"
        assert configs[1].xl_params.get("ylabel") == "Rental Count"
        assert configs[1].sql == """-- chart:line, xlabel=Date, ylabel=Rental Count
    SELECT rental_date :: date AS "Rental Date", COUNT(*) AS "Rental Count"
    FROM rental
    GROUP BY 1
    ORDER BY 1;"""

    os.remove(temp_file.name)


@pytest.mark.parametrize(
    "value, expected",
    [
        ("False", False),
        ("TRUE", True),
        ("0", 0),
        ("-42", -42),
        ("-3.14", -3.14),
        ("", ""),
        (3.14, 3.14),
        (True, True),
        (False, False),
        (0, 0),
    ],
)
def test_convert_edge_cases(value, expected):
    assert _convert(value) == expected


@pytest.mark.parametrize(
    "text, expected",
    [
        ("[1]", [1]),
        ("[1, 'two', 3.0]", [1, "two", 3.0]),
        ("[ 1 , 2 , 3 ]", [1, 2, 3]),
        ("(1,)", (1,)),
        ("[1,2,3]", [1, 2, 3]),
    ],
)
def test_parse_list_or_tuple_edge_cases(text, expected):
    assert _parse_list_or_tuple(text) == expected


def test_parse_sql_file_empty():
    sql_content = ";"
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 0
    os.remove(temp_file.name)


def test_parse_sql_file_chart_without_value():
    sql_content = """
    -- chart
    SELECT * FROM table1;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 1
        assert configs[0].xl_params.get("chart") == "chart"
    os.remove(temp_file.name)


def test_parse_sql_file_exec_directive():
    sql_content = """
    -- exec
    SELECT * FROM table1;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 1
        assert configs[0].xl_params.get("exec") == "chart"
    os.remove(temp_file.name)


def test_parse_sql_file_exec_with_value():
    sql_content = """
    -- exec=True
    SELECT * FROM table1;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 1
        assert configs[0].xl_params.get("exec") is True
    os.remove(temp_file.name)


def test_parse_sql_file_mixed_separators():
    sql_content = """
    -- chart:bar, title=My Chart, show_legend=false
    SELECT * FROM sales;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 1
        assert configs[0].xl_params.get("chart") == "bar"
        assert configs[0].xl_params.get("title") == "My Chart"
        assert configs[0].xl_params.get("show_legend") is False
    os.remove(temp_file.name)


def test_parse_sql_file_sheet_param():
    sql_content = """
    -- sheet_param country = US | UK
    SELECT * FROM sales WHERE country = :country;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 2
        assert configs[0].sql_params == {"country": "US"}
        assert configs[1].sql_params == {"country": "UK"}
        assert configs[0].xl_params.get("sheetname") == "US"
        assert configs[1].xl_params.get("sheetname") == "UK"
    os.remove(temp_file.name)


def test_parse_sql_file_sheet_param_single_value():
    sql_content = """
    -- sheet_param country = US
    SELECT * FROM sales WHERE country = :country;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 1
        assert configs[0].sql_params == {"country": "US"}
        assert configs[0].xl_params.get("sheetname") == "US"
    os.remove(temp_file.name)


def test_parse_sql_file_sheet_param_section_heading():
    sql_content = """
    -- sheet_param country = US | UK
    -- section_heading = Sales in country
    SELECT * FROM sales WHERE country = :country;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 2
        assert configs[0].xl_params.get("section_heading") == "Sales in US"
        assert configs[1].xl_params.get("section_heading") == "Sales in UK"
    os.remove(temp_file.name)


def test_parse_sql_file_sheet_param_duplicate_names():
    sql_content = """
    -- sheet_param country = US | UK
    -- sheet_param country = FR | DE
    SELECT * FROM sales;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        with pytest.raises(ValueError, match="Duplicate sheet_param"):
            parse_sql_file(temp_file.name)
    os.remove(temp_file.name)


def test_parse_sql_file_multiple_sheet_param_names():
    sql_content = """
    -- sheet_param country = US | UK
    -- sheet_param year = 2020 | 2021
    SELECT * FROM sales;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        with pytest.raises(ValueError, match="Multiple `sheet_param` names"):
            parse_sql_file(temp_file.name)
    os.remove(temp_file.name)


def test_parse_sql_file_query_param():
    sql_content = """
    -- query_param year = 2020 | 2021
    SELECT * FROM sales WHERE year = :year;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 2
        assert configs[0].sql_params == {"year": 2020}
        assert configs[1].sql_params == {"year": 2021}
    os.remove(temp_file.name)


def test_parse_sql_file_query_param_section_heading():
    sql_content = """
    -- query_param year = 2020 | 2021
    -- section_heading = Sales in year
    SELECT * FROM sales WHERE year = :year;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 2
        assert configs[0].xl_params.get("section_heading") == "Sales in 2020"
        assert configs[1].xl_params.get("section_heading") == "Sales in 2021"
    os.remove(temp_file.name)


def test_parse_sql_file_query_param_cartesian_product():
    sql_content = """
    -- query_param year = 2020 | 2021
    -- query_param region = East | West
    SELECT * FROM sales WHERE year = :year AND region = :region;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 4
        params = [c.sql_params for c in configs]
        assert {"year": 2020, "region": "East"} in params
        assert {"year": 2020, "region": "West"} in params
        assert {"year": 2021, "region": "East"} in params
        assert {"year": 2021, "region": "West"} in params
    os.remove(temp_file.name)


def test_parse_sql_file_sheet_and_query_param_combined():
    sql_content = """
    -- sheet_param country = US | UK
    -- query_param year = 2020 | 2021
    SELECT * FROM sales WHERE country = :country AND year = :year;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 4
        for config in configs:
            assert "country" in config.sql_params
            assert "year" in config.sql_params
    os.remove(temp_file.name)


def test_parse_sql_file_data_columns():
    sql_content = """
    -- chart:bar, data_columns = [1, 2, 3]
    SELECT * FROM sales;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 1
        assert configs[0].xl_params.get("data_columns") == [1, 2, 3]
    os.remove(temp_file.name)


def test_parse_sql_file_headings():
    sql_content = """
    -- chart:bar, headings = [Sales, Revenue, Profit]
    SELECT * FROM sales;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 1
        assert configs[0].xl_params.get("headings") == ["Sales", "Revenue", "Profit"]
    os.remove(temp_file.name)


def test_parse_sql_file_sheet_param_overrides_sheetname():
    sql_content = """
    -- sheet_param country = US | UK
    -- sheetname = CustomSheet
    SELECT * FROM sales WHERE country = :country;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 2
        assert configs[0].xl_params.get("sheetname") == "CustomSheet"
        assert configs[1].xl_params.get("sheetname") == "CustomSheet"
    os.remove(temp_file.name)


def test_parse_sql_file_irrelevant_comments_skipped():
    sql_content = """
    -- This is a regular comment
    -- chart:bar
    SELECT * FROM sales;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 1
        assert configs[0].xl_params.get("chart") == "bar"
    os.remove(temp_file.name)


def test_parse_sql_file_from_sql_script_flag():
    sql_content = """
    -- chart:bar
    SELECT * FROM sales;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 1
        assert configs[0].from_sql_script is True
    os.remove(temp_file.name)


def test_parse_sql_file_query_param_with_heading_replacement():
    sql_content = """
    -- query_param dept = HR | Engineering
    -- section_heading = Employees in dept
    SELECT * FROM employees WHERE dept = :dept;
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as temp_file:
        temp_file.write(sql_content.encode())
        temp_file.flush()
        configs = parse_sql_file(temp_file.name)
        assert len(configs) == 2
        assert configs[0].xl_params.get("section_heading") == "Employees in HR"
        assert configs[1].xl_params.get("section_heading") == "Employees in Engineering"
    os.remove(temp_file.name)
