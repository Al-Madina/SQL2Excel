"""
SQL2Excel — Generate Excel reports with charts from SQL queries and Pandas DataFrames.

Public API
----------
Report
    Top-level orchestrator that executes SQL queries and writes Excel workbooks.
QueryConfig
    Configuration object pairing a SQL query with Excel/chart parameters.
SQLExecutor
    Executes SQL queries via SQLAlchemy and returns pandas DataFrames.
parse_sql_file
    Parse an annotated ``.sql`` file into a list of ``QueryConfig`` objects.
Config
    Default configuration constants for charts, formatting, and layout.
ExcelHelper
    openpyxl utility layer for cell formatting, chart styling, and layout.

Chart Classes
-------------
Chart
    Base class for all chart types.
LineChart
    Line chart.
BarChart
    Bar chart.
PieChart
    Pie chart with data-label customization.
RadarChart
    Radar (spider) chart.
AreaChart
    Area chart.
ScatterChart
    Scatter plot with flexible data selection.
BubbleChart
    Bubble chart using x, y, and size columns.
StackedBarChart
    Stacked or percent-stacked bar chart.
BarLineChart
    Combined bar and line chart with dual y-axes.
SingleAxisBarLineChart
    Combined bar and line chart on a single y-axis.
ImageChart
    Embed images or matplotlib figures into Excel.
LineBasedChart
    Abstract base class for line and radar charts.
TwoAxesChart
    Abstract base class for charts with two y-axes.
"""

__version__ = "v0.1.0"

from sql2excel.chart import (
    AreaChart,
    BarChart,
    BarLineChart,
    BubbleChart,
    Chart,
    ImageChart,
    LineBasedChart,
    LineChart,
    PieChart,
    RadarChart,
    ScatterChart,
    SingleAxisBarLineChart,
    StackedBarChart,
    TwoAxesChart,
)
from sql2excel.config import Config
from sql2excel.excel_helper import ExcelHelper
from sql2excel.parser import parse_sql_file
from sql2excel.report import Report
from sql2excel.sqlexec import QueryConfig, SQLExecutor

__all__ = [
    "Report",
    "QueryConfig",
    "SQLExecutor",
    "parse_sql_file",
    "Config",
    "ExcelHelper",
    "Chart",
    "LineBasedChart",
    "LineChart",
    "BarChart",
    "PieChart",
    "RadarChart",
    "AreaChart",
    "ScatterChart",
    "BubbleChart",
    "StackedBarChart",
    "BarLineChart",
    "SingleAxisBarLineChart",
    "ImageChart",
    "TwoAxesChart",
]
