# Example of different line charts. To complete ...

import os

import openpyxl as xl
import pandas as pd

from sql2excel.chart import *
from sql2excel.config import Config


# Create a workbook
wb = xl.Workbook()
ws = wb.active


n_rows = 12
months = np.arange("2000-01", "2001-01", dtype="datetime64[M]")
# Format dates before sending them to the chart.
# NOTE: The current version does not support date axis formatting.
dates = np.array([np.datetime_as_string(month, unit="D") for month in months])
desktop = np.array([100 * (np.random.rand() * i / 2 + 1) for i in range(n_rows)])
laptop = np.array([300 * (np.random.rand() * i / 3 + 1) for i in range(n_rows)])
tablet = np.array([500 * (np.random.rand() * i / 4 + 1) for i in range(n_rows)])
phone = np.array([1000 * (np.random.rand() * i / 5 + 1) for i in range(n_rows)])

df = pd.DataFrame(
    {
        "Date": dates,
        "Desktop": desktop,
        "Laptop": laptop,
        "Baseline": [np.mean(laptop)] * len(laptop),
        "Tablet": tablet,
        "Phone": phone,
    }
)

# NOTE: The baseline is detected and formatted differently
# A baseline is a reference line with a constant value.

# Basic plot
chart = LineChart()
chart.plot(df, ws)

# Use the default openpyxl colors
chart = LineChart()
chart.plot(df, ws, openpyxl_color=True, section_heading="Using Openpyxl colors")

# Not smooth
chart = LineChart()
chart.plot(df, ws, smooth=False, section_heading="Not smooth line")

# Not using the default customization of the reference line
chart = LineChart()
chart.plot(
    df,
    ws,
    use_ref_line=False,
    section_heading="Not using the default customization of the basline (reference line)",
)

# Not using a reference line by overriding the default config
config = Config()
# Override the default
config.USE_REF_LINE = False
# Pass the config object to the Chart
chart = LineChart(config=config)
chart.plot(
    df,
    ws,
    section_heading="Not using the default customization of the basline (reference line) by overriding default config",
)


# Some customizations
chart = LineChart()
chart.plot(
    df,
    ws,
    section_heading="Monthly sales",
    title="Monthly sales",
    line_width=2,
    line_style="sysDash",
    # line_color="FF0000", Do not use this unless you have one line
    marker_symbol="circle",
    marker_size=8,
    # row_start=40,
    xlabel="Month",
    ylabel="Sales",
    width=24,  # Chart width
    height=12,  # Chart height
    chart_position="bottom",
    rotation=90,  # x-axis major ticks
    # show_legend=False,  # You can disable legend
    legend_position="t",  # legend position: 'r', 't', 'l', 'b'
    # openpyxl_color=True,
)


file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "line_chart.xlsx"
)
wb.save(file_name)
