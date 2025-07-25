"""
Illustrate Pie chart options
"""

import os

from numpy.random import randint
from openpyxl import Workbook
from pandas import DataFrame

from sql2excel.chart import PieChart

n_slices = 10

fields = [f"Scientific Field ({i})" for i in range(1, n_slices + 1)]

values = randint(100, 1000, size=n_slices)

df = DataFrame({"Field": fields, "Publication": values})

wb = Workbook()
ws = wb.active

chart = PieChart()
chart.plot(df, ws, section_heading="Pie Chart: default options")

# All options are turned off
chart.plot(
    df,
    ws,
    section_heading="Pie Chart: all options are False",
    show_percentage=False,
    show_category=False,
    show_legend_key=False,
    show_values=False,
    show_series_name=False,
    show_legend=False,
)

# Showing percentages and values
chart.plot(
    df,
    ws,
    section_heading="Pie Chart: showing percentages and values",
    show_percentage=True,
    show_category=False,
    show_legend_key=False,
    show_values=True,
    show_series_name=False,
)

# Showing categories, percentages, and values
chart.plot(
    df,
    ws,
    section_heading="Pie Chart: showing categories, percentages, and values",
    show_percentage=True,
    show_category=True,
    show_legend_key=False,
    show_values=True,
    show_series_name=False,
)

# Showing everything
chart.plot(
    df,
    ws,
    section_heading="Pie Chart: showing everything",
    show_percentage=True,
    show_category=True,
    show_legend_key=True,
    show_values=True,
    show_series_name=True,
    title="This is a custom title",
)

# Using Openpyxl default colors
chart.plot(
    df,
    ws,
    section_heading="Pie Chart: openpyxl colors",
    show_percentage=True,
    show_category=False,
    show_legend_key=False,
    show_values=True,
    show_series_name=False,
    openpyxl_color=True,
)


file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "pie_chart.xlsx"
)

wb.save(file_name)
