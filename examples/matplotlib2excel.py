"""

"""

import os

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import Workbook

from sql2excel.chart import ImageChart

# Create workbook
wb = Workbook()
ws = wb.active

# To insert the figure, use ImageChart
chart = ImageChart()


# Dummy data
df = pd.DataFrame({"X": np.linspace(0, 1, 11), "Y": np.random.uniform(size=11)})

# Using plt
fig = plt.figure()
plt.plot(df.X, df.Y)
plt.xlabel("X")
plt.ylabel("Y")
plt.title("Using plt")
plt.close(fig)

# Default insertion
chart.add_image(fig, ws, section_heading="Inserting Matplotlib figures (using plt)")


# Using axes
fig, ax = plt.subplots()
ax.plot(df.X, df.Y)
plt.xlabel("X")
plt.ylabel("Y")
plt.title("Using Axes")
plt.close(fig)

# Default insertion again
chart.add_image(fig, ws, section_heading="Inserting Matplotlib figures (using axes)")

# When you pass dataframe, it will be written as well
chart.add_image(
    fig,
    ws,
    df=df,
    section_heading="When you pass dataframe, it will be written as well",
)

# Same as above, but write the figure at the bottom
chart.add_image(
    fig,
    ws,
    df=df,
    section_heading="Passing dataframe and specifying figure position relative to dataframe",
    chart_position="bottom",
)

# One final insertion!
chart.add_image(
    fig,
    ws,
    section_heading="Final insertion to show the spacing is calculated correctly",
)

file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "matplotlib2excel.xlsx"
)

# Save the workbook
wb.save(file_name)
