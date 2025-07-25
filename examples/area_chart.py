# An example of different area charts
import os

import openpyxl as xl
import pandas as pd

from sql2excel.chart import *

wb = xl.Workbook()
ws = wb.active

rows = [
    ["JAN", 40, 30],
    ["FEB", 40, 25],
    ["MAR", 50, 30],
    ["APR", 30, 10],
    ["MAY", 25, 5],
    ["JUN", 50, 10],
]

df = pd.DataFrame(data=rows, columns=["Month", "Batch 1", "Batch 2"])
df = df[["Month", "Batch 2", "Batch 1"]]
chart = AreaChart()
chart.plot(
    df,
    ws,
    section_heading="Area Chart",
    show_legend=True,
    # shape_line_color="FFFFFF",
    chart_position="bottom",
)

file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "area_chart.xlsx"
)
wb.save(file_name)
