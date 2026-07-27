import pytest

from sql2excel.chart import (
    AreaChart,
    BarChart,
    BarLineChart,
    BubbleChart,
    Chart,
    LineChart,
    PieChart,
    RadarChart,
    ScatterChart,
    SingleAxisBarLineChart,
    StackedBarChart,
)
from sql2excel.report import CHART_MAP, Report
from sql2excel.sqlexec import QueryConfig

from .setup import *


def test_chart_map_contains_expected_keys():
    expected = {
        "chart",
        "area",
        "line",
        "bar",
        "barline",
        "pie",
        "singleaxisbarline",
        "radar",
        "bubble",
        "scatter",
        "stackedbar",
    }
    assert set(CHART_MAP.keys()) == expected


def test_chart_map_values_are_correct_classes():
    assert CHART_MAP["chart"] is Chart
    assert CHART_MAP["line"] is LineChart
    assert CHART_MAP["bar"] is BarChart
    assert CHART_MAP["pie"] is PieChart
    assert CHART_MAP["area"] is AreaChart
    assert CHART_MAP["radar"] is RadarChart
    assert CHART_MAP["scatter"] is ScatterChart
    assert CHART_MAP["bubble"] is BubbleChart
    assert CHART_MAP["stackedbar"] is StackedBarChart
    assert CHART_MAP["barline"] is BarLineChart
    assert CHART_MAP["singleaxisbarline"] is SingleAxisBarLineChart


def test_report_default_init(sqlite_engine):
    report = Report(engine=sqlite_engine)
    assert report.config is not None
    assert report.excel_helper is not None
    assert report.executor is not None
    report.close()


def test_report_init_with_engine(sqlite_engine):
    report = Report(engine=sqlite_engine)
    assert report.executor.engine is sqlite_engine
    report.close()


def test_report_init_with_config():
    from sql2excel.config import Config

    cfg = Config()
    report = Report(engine=create_engine("sqlite:///:memory:"), config=cfg)
    assert report.config is cfg
    report.close()


def test_report_init_silent_flag(tmp_path, sqlite_engine):
    report = Report(engine=sqlite_engine, silent=True)
    assert report.executor.silent is True

    # Silent mode will not throw an exception but warning
    with pytest.warns(UserWarning, match="Unable to execute the query"):
        qc = QueryConfig(sql="SELECT * FROM non_existent_table")
        out = tmp_path / "out.xlsx"
        report.generate(qc, file_name=str(out))

    report.close()


def test_generate_single_query_without_chart(tmp_path, sqlite_engine):
    out = tmp_path / "out.xlsx"
    qc = QueryConfig(sql="SELECT name, salary FROM employees ORDER BY id")

    report = Report(engine=sqlite_engine)
    report.generate(qc, file_name=str(out))
    report.close()

    assert out.exists()
    wb = xl.load_workbook(str(out))
    ws = wb.active
    assert ws.cell(1, 1).value is None
    wb.close()


def test_generate_single_query_with_chart(tmp_path, sqlite_engine):
    out = tmp_path / "out.xlsx"
    qc = QueryConfig(
        sql="SELECT name, salary FROM employees ORDER BY id",
        chart="chart",
    )

    report = Report(engine=sqlite_engine)
    report.generate(qc, file_name=str(out))
    report.close()

    wb = xl.load_workbook(str(out))
    ws = wb.active
    assert ws.cell(1, 1).value == "name"
    assert ws.cell(1, 2).value == "salary"
    assert ws.cell(2, 1).value == "Alice"
    wb.close()


def test_generate_accepts_list(tmp_path, sqlite_engine):
    out = tmp_path / "out.xlsx"
    qc1 = QueryConfig(
        sql="SELECT name FROM employees WHERE name = 'Alice'", chart="chart"
    )
    qc2 = QueryConfig(
        sql="SELECT name FROM employees WHERE name = 'Bob'", chart="chart"
    )

    report = Report(engine=sqlite_engine)
    report.generate([qc1, qc2], file_name=str(out))
    report.close()

    wb = xl.load_workbook(str(out))
    ws = wb.active
    assert ws.cell(1, 1).value == "name"
    assert ws.cell(2, 1).value == "Alice"

    # Bob can be anywhere below depending on how much space the user set on config
    is_bob_found = False
    for cell_value in ws.iter_rows(
        min_row=1, max_row=10, min_col=1, max_col=1, values_only=True
    ):
        is_bob_found = cell_value[0] == "Bob"
        if is_bob_found:
            break

    assert is_bob_found

    wb.close()


def test_generate_sheetname_kwarg(tmp_path, sqlite_engine):
    out = tmp_path / "out.xlsx"
    qc = QueryConfig(sql="SELECT name FROM employees")

    report = Report(engine=sqlite_engine)
    report.generate(qc, file_name=str(out), sheetname="MySheet")
    report.close()

    wb = xl.load_workbook(str(out))
    assert "MySheet" in wb.sheetnames
    wb.close()


def test_generate_sheetname_from_xl_params(tmp_path, sqlite_engine):
    out = tmp_path / "out.xlsx"
    qc = QueryConfig(
        sql="SELECT name FROM employees",
        sheetname="XlSheet",
    )

    report = Report(engine=sqlite_engine)
    report.generate(qc, file_name=str(out))
    report.close()

    wb = xl.load_workbook(str(out))
    assert "XlSheet" in wb.sheetnames
    wb.close()


def test_generate_default_filename(tmp_path, monkeypatch, sqlite_engine):
    monkeypatch.chdir(tmp_path)
    qc = QueryConfig(sql="SELECT name FROM employees")

    report = Report(engine=sqlite_engine)
    report.generate(qc)
    report.close()

    assert (tmp_path / "result.xlsx").exists()


def test_generate_skips_query_without_chart_or_exec(tmp_path, sqlite_engine):
    out = tmp_path / "out.xlsx"
    qc = QueryConfig(sql="SELECT name FROM employees")

    report = Report(engine=sqlite_engine)
    report.generate(qc, file_name=str(out))
    report.close()

    wb = xl.load_workbook(str(out))
    ws = wb.active
    assert ws.cell(1, 1).value is None
    wb.close()


def test_generate_skips_exec_in_xl_params(tmp_path, sqlite_engine):
    out = tmp_path / "out.xlsx"
    qc = QueryConfig(
        sql="SELECT name FROM employees",
        exec=True,
    )

    report = Report(engine=sqlite_engine)
    report.generate(qc, file_name=str(out))
    report.close()

    wb = xl.load_workbook(str(out))
    ws = wb.active
    assert ws.cell(1, 1).value is None
    wb.close()


def test_generate_with_chart(tmp_path, sqlite_engine):
    out = tmp_path / "out.xlsx"
    qc = QueryConfig(
        sql="SELECT name, salary FROM employees ORDER BY id",
        chart="chart",  # Any chart will do
        section_heading="Employee Salaries",
        title="Salary Chart",
    )

    report = Report(engine=sqlite_engine)
    report.generate(qc, file_name=str(out))
    report.close()

    assert out.exists()

    wb = xl.load_workbook(str(out))
    ws = wb.active
    assert ws.cell(1, 1).value == "Employee Salaries"
    assert ws.cell(2, 1).value == "name"
    assert ws.cell(2, 2).value == "salary"
    wb.close()


def test_generate_with_line_chart(tmp_path, sqlite_engine):
    out = tmp_path / "out.xlsx"
    qc = QueryConfig(
        sql="SELECT name, salary FROM employees ORDER BY id",
        chart="line",
        title="Line Chart",
    )

    report = Report(engine=sqlite_engine)
    report.generate(qc, file_name=str(out))
    report.close()

    assert out.exists()


def test_generate_with_pivot(tmp_path, sqlite_engine):
    out = tmp_path / "out.xlsx"
    qc = QueryConfig(
        sql="SELECT department, name, salary FROM employees ORDER BY department, name",
        chart="chart",
        index="department",
        columns="name",
        values="salary",
    )

    report = Report(engine=sqlite_engine)
    report.generate(qc, file_name=str(out))
    report.close()

    wb = xl.load_workbook(str(out))
    ws = wb.active

    # The employee name will be the columns if the data is pivoted
    assert ws.cell(1, 1).value == "department"
    assert ws.cell(1, 2).value == "Alice"
    assert ws.cell(1, 3).value == "Bob"
    assert ws.cell(1, 4).value == "Charlie"

    wb.close()

    assert out.exists()


def test_close_sets_executor_closed(sqlite_engine):
    report = Report(engine=sqlite_engine)
    assert report.executor.closed is False
    report.close()
    assert report.executor.closed is True


def test_generate_closes_executor_on_exception(tmp_path, sqlite_engine):
    out = tmp_path / "out.xlsx"
    qc = QueryConfig(sql="INVALID SQL ???")

    report = Report(engine=sqlite_engine)
    with pytest.raises(Exception):
        report.generate(qc, file_name=str(out))
    assert report.executor.closed is True
