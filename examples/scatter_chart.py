"""
Illustrate Scatter chart options
"""

import os

import numpy as np
from openpyxl import Workbook
from pandas import DataFrame

from sql2excel.chart import ScatterChart

n_points = 15

np.random.seed(42)
x = np.linspace(0, 10, n_points)
y1 = 2 * x + np.random.normal(0, 1.5, n_points)
y2 = -x + np.random.normal(5, 2, n_points)
y3 = 0.5 * x**2 + np.random.normal(0, 3, n_points)

df = DataFrame({"X": np.round(x, 2), "Series A": np.round(y1, 2), "Series B": np.round(y2, 2), "Series C": np.round(y3, 2)})

wb = Workbook()
ws = wb.active

# Default scatter chart
chart = ScatterChart()
chart.plot(df, ws, section_heading="Scatter Chart: default options")

# Custom title and axis labels
chart = ScatterChart()
chart.plot(
    df,
    ws,
    section_heading="Scatter Chart: custom title and labels",
    title="X vs Y Values",
    xlabel="X axis",
    ylabel="Y axis",
)

# Custom marker symbols and size
chart = ScatterChart()
chart.plot(
    df,
    ws,
    section_heading="Scatter Chart: custom marker symbols and size",
    marker_symbols=["circle", "triangle", "star"],
    marker_size=10,
)

# Non-default chart style
chart = ScatterChart()
chart.plot(
    df,
    ws,
    section_heading="Scatter Chart: custom chart style",
    chart_style=2,
)

# Using custom colors
chart = ScatterChart()
chart.plot(
    df,
    ws,
    section_heading="Scatter Chart: custom openpyxl colors",
    custom_colors=True,
)

# With lines connecting markers (nofill=False)
chart = ScatterChart()
chart.plot(
    df,
    ws,
    section_heading="Scatter Chart: lines connecting markers",
    nofill=False,
)

# Custom dimensions and legend position
chart = ScatterChart()
chart.plot(
    df,
    ws,
    section_heading="Scatter Chart: custom dimensions and legend",
    width=20,
    height=10,
    legend_position="r",
)

# With rotation on x-axis tick labels
chart = ScatterChart()
chart.plot(
    df,
    ws,
    section_heading="Scatter Chart: rotated x-axis labels",
    rotation=45,
)

# Select specific data columns
chart = ScatterChart()
chart.plot(
    df,
    ws,
    section_heading="Scatter Chart: specific data columns",
    data_columns=[2, 4],
)

# With custom titles for each heading
chart = ScatterChart()
chart.plot(
    df,
    ws,
    section_heading="Scatter Chart: everything customized",
    title="Custom Scatter Plot",
    xlabel="Input",
    ylabel="Output",
    marker_symbols=["diamond", "x", "square"],
    marker_size=8,
    width=22,
    height=11,
    legend_position="t",
    custom_colors=True,
    chart_position="bottom",
)

file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "scatter_chart.xlsx"
)

wb.save(file_name)
