"""
An example to show how to insert an image given its path
"""

import os

import pandas as pd
from openpyxl import Workbook

from sql2excel.chart import ImageChart

# Create workbook
wb = Workbook()
ws = wb.active

# Read data
df = pd.read_csv(
    "data/Africa_collab_2013_2022.csv", delimiter=";", keep_default_na=False
)

# Insert into excel
chart = ImageChart()
# The image is very large. Adjust width and height while preserving aspect ratio
# By default, the image is inserted next to the data (on the right)
chart.add_image("data/Africa_collab_2013_2022.png", ws, df=df, width=1425, height=552)

# You can also insert the image below the data
chart.add_image(
    "data/Africa_collab_2013_2022.png",
    ws,
    df=df,
    width=1425,
    height=552,
    chart_position="bottom",
)


file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "image2excel.xlsx"
)

# Save excel file
wb.save(file_name)
