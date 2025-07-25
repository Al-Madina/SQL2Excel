import os

import openpyxl as xl
import pandas as pd

from sql2excel.chart import BubbleChart

fields = [
    "Agricultural Sciences",
    "Engineering",
    "Health Sciences",
    "Humanities",
    "Natural Sciences",
    "Social Sciences",
]

rfs = [1.3, 0.4, 1.5, 1.9, 0.7, 1.1]
mncs = [0.9, 1.1, 1.8, 0.6, 1.6, 0.5]
npubs = [200, 500, 1000, 150, 1200, 700]

df = pd.DataFrame(
    {
        "Field": fields,
        "Relative Field Strength": rfs,
        "Mean-Normalized Citation Score": mncs,
        "Number of Publications": npubs,
    }
)


wb = xl.Workbook()
ws = wb.active

chart = BubbleChart()
chart.plot(
    df,
    ws,
    section_heading="Bubble Chart: positional analysis",
    xlabel="Relative Field Strength",
    ylabel="Mean-Normalized Citation Score",
)


chart.plot(
    df,
    ws,
    section_heading="Bubble Chart (Openpyxl colors): positional analysis",
    xlabel="Relative Field Strength",
    ylabel="Mean-Normalized Citation Score",
    openpyxl_color=True,
)

file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "bubble_chart.xlsx"
)
wb.save(file_name)
