"""Examples of stacked bar charts"""

import os

import numpy as np
import openpyxl as xl
import pandas as pd

from sql2excel.chart import StackedBarChart

fields = [
    "Agricultural Sciences",
    "Engineering",
    "Health Sciences",
    "Humanities",
    "Natural Sciences",
    "Social Sciences",
]

pubyear = np.arange(2020, 2025)


data = []

# Generate random values for each field for each publication year
for year in pubyear:
    for field in fields:
        value = np.random.randint(100, 1000)
        data.append([year, field, value])

# Create the DataFrame
df = pd.DataFrame(data, columns=["pubyear", "field", "value"])

wb = xl.Workbook()
ws = wb.active

chart = StackedBarChart()

chart.write_dataframe(
    df, ws, section_heading="Original data is not suitable. Need to pivot it"
)

# Pivot data
df1 = pd.pivot(df, index="pubyear", columns="field", values="value").reset_index(
    drop=False
)

chart.plot(
    df1,
    ws,
    section_heading="Adding data from columns: Stackedbar chart after pivoting data",
)


df2 = pd.pivot(df, index="field", columns="pubyear", values="value").reset_index(
    drop=False
)

chart.plot(
    df2,
    ws,
    from_rows=True,  # Use this option if your data is arranged in rows
    section_heading="Adding data from rows: Stakcedbar chart after pivoting data differently",
)

chart.plot(
    df1,
    ws,
    chart_grouping="stacked",
    section_heading="Stacked (not adding up to 100%): Stakcedbar chart after pivoting data differently",
)

chart.plot(
    df1,
    ws,
    openpyxl_color=True,
    section_heading="Using default Openpyxl colors: Stackedbar chart after pivoting data",
)

file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "stackedbar_chart.xlsx"
)
wb.save(file_name)
