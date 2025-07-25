from setup import *

from sql2excel.sqlexec import (
    _bind_positional_parameters,
    _convert_query_placeholders,
)


@pytest.mark.parametrize(
    "query, params, expected_query",
    [
        (
            "SELECT * FROM table WHERE id = ?",
            [1],
            "SELECT * FROM table WHERE id = :param_1",
        ),
        (
            "SELECT * FROM table WHERE id = ? AND name = ?",
            [1, "John"],
            "SELECT * FROM table WHERE id = :param_1 AND name = :param_2",
        ),
        ("SELECT * FROM table", [], "SELECT * FROM table"),
        (
            "INSERT INTO table (id, name) VALUES (?, ?)",
            [1, "Alice"],
            "INSERT INTO table (id, name) VALUES (:param_1, :param_2)",
        ),
    ],
)
def test_convert_query_placeholders(query, params, expected_query):
    assert _convert_query_placeholders(query, params) == expected_query


# Test cases for _bind_positional_parameters
@pytest.mark.parametrize(
    "query, params, expected_query, expected_params",
    [
        (
            "SELECT * FROM table WHERE id = ?",
            [1],
            "SELECT * FROM table WHERE id = :param_1",
            {"param_1": 1},
        ),
        (
            "SELECT * FROM table WHERE id = ? AND name = ?",
            [1, "John"],
            "SELECT * FROM table WHERE id = :param_1 AND name = :param_2",
            {"param_1": 1, "param_2": "John"},
        ),
        (
            "INSERT INTO table (id, name) VALUES (?, ?)",
            [1, "Alice"],
            "INSERT INTO table (id, name) VALUES (:param_1, :param_2)",
            {"param_1": 1, "param_2": "Alice"},
        ),
        (
            "SELECT * FROM table WHERE id IN ?",
            [[1, 2, 3]],
            "SELECT * FROM table WHERE id IN :param_1",
            {"param_1": [1, 2, 3]},
        ),
    ],
)
def test_bind_positional_parameters(query, params, expected_query, expected_params):
    bound_sql, bound_params = _bind_positional_parameters(query, params)
    assert str(bound_sql.text) == expected_query
    assert bound_params == expected_params
    for key, value in expected_params.items():
        print(bound_sql._bindparams[key])
        assert bound_sql._bindparams[key].key == key
        assert bound_sql._bindparams[key].value == value
