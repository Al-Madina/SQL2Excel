# An example of different barline chart
import os

import numpy as np
import openpyxl as xl
import pandas as pd

from sql2excel.chart import *

wb = xl.Workbook()
ws = wb.active

n_rows = 11

df = pd.DataFrame(
    {
        "pubyear": np.arange(2000, 2000 + n_rows),
        "pubcount": [100 * (i / 2 + 1) for i in range(n_rows)],
        "share": [0.1 + (i * np.random.rand()) / 20 for i in range(n_rows)],
    }
)


# Basic plot
chart = BarLineChart()
chart.plot(df, ws)

# Some customizations
chart = BarLineChart()
chart.plot(
    df,
    ws,
    section_heading="Barline chart with some customization applied differently to the two axes",
    title="Publication count and world share",
    # You can provide custom headings for the columns
    headings=["Publication Year", "Article Count", "World Share"],
    line_width=3,
    line_style="solid",
    line_color="107C10",
    smooth=False,
    marker_symbol="square",
    marker_size=8,
    column_start=3,
    ylabel1="Publication count",
    ylabel2="World share",
    ylim1=(0, 1000),
    ylim2=(0, 1),
    y_orientation2="maxMin",
    width=20,
    height=10,
    chart_position="bottom",
    rotation=0,
    axis_bold_font=True,
)

# More columns
df["share2"] = 1.5 * df["share"]
df["share3"] = 1.9 * df["share"]

chart = BarLineChart()
chart.plot(df, ws)


# With Openpyxl default colors
chart = BarLineChart()
chart.plot(df, ws, openpyxl_color=True)

# More columns with `data_columns` to select a few
chart = BarLineChart()
chart.plot(df, ws, data_columns=[2, 3])


file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "barline_chart.xlsx"
)
wb.save(file_name)
