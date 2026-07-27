from sql2excel.sqlexec import (
    QueryConfig,
    _bind_positional_parameters,
    _convert_query_placeholders,
)

from .setup import *


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


def test_execute_simple_query(executor):
    """Test standard SELECT returning a pandas DataFrame."""
    qc = QueryConfig(sql="SELECT name, salary FROM employees ORDER BY id")
    df = executor.execute(qc)

    assert isinstance(df, pd.DataFrame)
    assert list(df.columns) == ["name", "salary"]
    assert len(df) >= 1
    assert df.iloc[0]["name"] == "Alice"


def test_execute_positional_parameters(sqlite_engine):
    """Test SQL with ? positional parameters and IN clauses (expanding params)."""
    exec_obj = SQLExecutor(engine=sqlite_engine)

    # 1. Single positional
    qc1 = QueryConfig(
        sql="SELECT * FROM employees WHERE department = ?", sql_params=("Engineering",)
    )
    df1 = exec_obj.execute(qc1)
    assert len(df1) == 2

    # 2. Expanding IN list positional
    qc2 = QueryConfig(
        sql="SELECT * FROM employees WHERE department IN ?",
        sql_params=[("Engineering", "HR")],
    )
    df2 = exec_obj.execute(qc2)
    assert len(df2) == 3
    exec_obj.close()


def test_execute_named_parameters(sqlite_engine):
    """Test SQL with named :param style."""
    exec_obj = SQLExecutor(engine=sqlite_engine)
    qc = QueryConfig(
        sql="SELECT * FROM employees WHERE salary > :min_sal",
        sql_params={"min_sal": 70000},
    )
    df = exec_obj.execute(qc)
    assert len(df) == 2
    exec_obj.close()


def test_execute_no_rows_returned(sqlite_engine):
    """Statements like CREATE/DROP/INSERT return None instead of crashing."""
    exec_obj = SQLExecutor(engine=sqlite_engine)
    qc = QueryConfig(sql="CREATE TABLE dummy (id INT)")
    df = executor_res = exec_obj.execute(qc)
    assert df is None
    exec_obj.close()


def test_silent_error_handling(sqlite_engine):
    """When silent=True, errors produce warnings and return None instead of raising."""
    exec_obj = SQLExecutor(engine=sqlite_engine, silent=True)
    qc = QueryConfig(sql="SELECT * FROM non_existent_table")

    with pytest.warns(UserWarning, match="Unable to execute the query"):
        df = exec_obj.execute(qc)

    assert df is None
    exec_obj.close()


def test_invalid_init_raises():
    """Initializing without any target raises ValueError."""
    with pytest.raises(ValueError, match="Cannot establish a connection"):
        SQLExecutor()
