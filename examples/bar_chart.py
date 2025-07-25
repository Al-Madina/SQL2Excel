import os

import numpy as np
import openpyxl as xll
import pandas as pd

from sql2excel.chart import BarChart

# Dummy data of electronic sales
n_rows = 12
months = np.arange("2000-01", "2001-01", dtype="datetime64[M]")
dates = np.array([np.datetime_as_string(month, unit="D") for month in months])
desktop = np.array([100 * (np.random.rand() * i / 2 + 1) for i in range(n_rows)])
laptop = np.array([300 * (np.random.rand() * i / 3 + 1) for i in range(n_rows)])
tablet = np.array([500 * (np.random.rand() * i / 4 + 1) for i in range(n_rows)])

df = pd.DataFrame(
    {
        "Date": dates,
        "Desktop": desktop,
        "Laptop": laptop,
        "Tablet": tablet,
    }
)

wb = xll.Workbook()
ws = wb.active

chart = BarChart()
chart.plot(df, ws, section_heading="Bar Chart: default (column type)")

chart.plot(df, ws, openpyxl_color=True, section_heading="Bar Chart: openpyxl colors")

chart.plot(
    df,
    ws,
    section_heading="Bar Chart: bar type",
    chart_type="bar",
    chart_position="bottom",
    legend_position="r",
    height=15,
)

# Negative bars
df = pd.DataFrame(
    {
        "x-axis": np.arange(1, 11),
        "y-axis": np.random.uniform(0, 1, size=10)
        - np.random.uniform(0.5, 0.5, size=10),
    }
)

chart = BarChart()
chart.plot(
    df,
    ws,
    section_heading="Negative bars: xlabel position at the bottom",
    xtick_label_position="low",
)

df["y-axis"] = 1
chart = BarChart()
chart.plot(
    df,
    ws,
    section_heading="Varying color for each bar",
    vary_color=True,
)


file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "bar_chart.xlsx"
)
wb.save(file_name)
