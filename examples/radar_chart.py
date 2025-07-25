"""Examples of different radar charts."""

import os

import openpyxl as xl
import pandas as pd

from sql2excel.chart import RadarChart


fields = [
    "Agricultural Sciences",
    "Engineering",
    "Health Sciences",
    "Humanities",
    "Natural Sciences",
    "Social Sciences",
]

data = {
    "Field": fields,
    "2003-2012": (1.2, 0.5, 1.5, 2, 0.9, 1.3),
    "2013-2022": (1.5, 0.95, 1.1, 1.8, 0.7, 1.2),
}

df = pd.DataFrame(data=data)

chart = RadarChart()


wb = xl.Workbook()
ws = wb.active

# Default
chart.plot(
    df,
    ws,
    section_heading="Default: Relative Field Strength - Dummy Data",
    width=12,
    height=12,
)


# Filled
chart.plot(
    df,
    ws,
    chart_type="filled",
    section_heading="Filled: Relative Field Strength - Dummy Data",
    width=12,
    height=12,
)


# Adding a reference (baseline)
average = pd.Series([1] * 6, name="Average")
df = pd.concat((df, average), axis=1)
chart.plot(
    df,
    ws,
    section_heading="With reference or baseline: Relative Field Strength - Dummy Data",
    width=12,
    height=12,
)

# Radar unit (y-axis major unit): you can also set the unit manually
chart.plot(
    df,
    ws,
    section_heading="Radar Unit (= 1): Relative Field Strength - Dummy Data",
    width=12,
    height=12,
    radar_unit=1,
)

chart.plot(
    df,
    ws,
    section_heading="Radar Unit (= 0.5): Relative Field Strength - Dummy Data",
    width=12,
    height=12,
    radar_unit=0.5,
)


# Radar unit steps: lower values lead to lower number of levels (rings)
chart.plot(
    df,
    ws,
    section_heading="Radar Unit Steps: Relative Field Strength - Dummy Data",
    width=12,
    height=12,
    radar_unit_steps=1.8,
)

# y-axis major unit = (max(df) - min(df)) / radar_unit_steps

# Yet another example with radar unit steps: lower values lead to lower number of levels (rings)
chart.plot(
    df,
    ws,
    section_heading="Yet another example with radar unit steps: Relative Field Strength - Dummy Data",
    width=12,
    height=12,
    radar_unit_steps=7,
)


file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "radar_chart.xlsx"
)
wb.save(file_name)
