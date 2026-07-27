import numpy as np
import openpyxl as xl
import pandas as pd
import pytest
from sqlalchemy import create_engine, text
from sqlalchemy.orm import Session

from sql2excel.chart import Chart
from sql2excel.config import Config
from sql2excel.excel_helper import ExcelHelper
from sql2excel.sqlexec import SQLExecutor

from .test_utils import *


@pytest.fixture(scope="session", autouse=True)
def workbook():
    wb = xl.Workbook()
    yield wb
    wb.close()


@pytest.fixture(scope="session", autouse=True)
def worksheet(workbook):
    ws = workbook.active
    return ws


@pytest.fixture(scope="session")
def config():
    return Config()


@pytest.fixture(scope="session", autouse=True)
def excel_helper(config):
    excel_helper = ExcelHelper(config=config)
    return excel_helper


@pytest.fixture(scope="session", autouse=True)
def chart(config, excel_helper):
    return Chart(config=config, excel_helper=excel_helper)


@pytest.fixture(scope="session", autouse=True)
def df():
    n_rows = 10
    data = {
        "pubyear": np.arange(2000, 2000 + n_rows),
        "pubcount": np.random.randint(100, 1000, size=n_rows),
        "perc": np.random.rand(n_rows),
    }

    return pd.DataFrame(data)


@pytest.fixture(scope="function")
def random_dataframes():
    def _generate_dataframes(n):
        dataframes = []
        for _ in range(n):
            rows = np.random.randint(3, 10)
            cols = np.random.randint(2, 5)
            df = pd.DataFrame(
                np.random.randint(0, 100, size=(rows, cols)),
                columns=[f"Col{i+1}" for i in range(cols)],
            )
            dataframes.append(df)
        return dataframes

    return _generate_dataframes


@pytest.fixture(scope="session", autouse=True)
def long_wide_df():
    n_rows = 10
    n_cols = 6
    data = np.random.rand(n_rows, n_cols - 1)
    columns = [f"Column_{i+1}" for i in range(n_cols - 1)]
    df = pd.DataFrame(data, columns=columns)
    df.insert(
        0,
        column="pubyear",
        value=np.arange(2000, 2000 + n_rows),
    )
    return df


@pytest.fixture(scope="function")
def create_worksheet(workbook):
    return workbook.create_sheet()


@pytest.fixture(scope="function")
def clear_worksheet(worksheet):
    while worksheet.max_row > 1:
        worksheet.delete_rows(2)

    worksheet.delete_rows(1)
    return worksheet


@pytest.fixture(scope="function")
def add_content_to_sheet(worksheet):
    n_rows = np.random.randint(1, 11)
    for r in range(n_rows):
        worksheet.append([f"Row #{r}"])

    return worksheet


@pytest.fixture
def sqlite_engine():
    """In-memory SQLite engine initialized with sample data."""
    engine = create_engine("sqlite:///:memory:")
    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE employees (
                id INTEGER PRIMARY KEY,
                name TEXT,
                department TEXT,
                salary REAL
            )
        """))
        conn.execute(text("""
            INSERT INTO employees (name, department, salary) VALUES
            ('Alice', 'Engineering', 90000),
            ('Bob', 'HR', 60000),
            ('Charlie', 'Engineering', 85000)
        """))
    return engine


# NOTE Run every test that uses this fixture once for every value.
@pytest.fixture(params=["db_url", "engine", "conn", "session"])
def executor(request, sqlite_engine, tmp_path):
    """Parametrized fixture testing all 4 ways to initialize SQLExecutor."""

    # mode gets its value from `params`
    mode = request.param

    if mode == "engine":
        exec_obj = SQLExecutor(engine=sqlite_engine)
    elif mode == "conn":
        conn = sqlite_engine.connect()
        exec_obj = SQLExecutor(conn=conn)
    elif mode == "session":
        session = Session(sqlite_engine)
        exec_obj = SQLExecutor(session=session)
    elif mode == "db_url":
        # Create a file-based SQLite database for URL testing
        db_path = tmp_path / "test.db"
        url = f"sqlite:///{db_path}"
        # Populate file DB
        init_engine = create_engine(url)
        with init_engine.begin() as conn:
            conn.execute(
                text("CREATE TABLE employees (id INT, name TEXT, salary REAL)")
            )
            conn.execute(text("INSERT INTO employees VALUES (1, 'Alice', 90000)"))
        exec_obj = SQLExecutor(db_url=url)

    yield exec_obj
    exec_obj.close()
