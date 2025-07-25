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

chart = Chart()
# Default options
chart.write_dataframe(df, ws, section_heading="Writing data without chart")

# Full options
chart = BarLineChart()
chart.plot(
    df,
    ws,
    section_heading="Publication count and world share",
    title="Publication count and world share",
    # You can provide custom headings for the columns
    headings=["Publication Year", "Article Count", "World Share"],
    line_width=3,
    line_style="solid",
    line_color="FD625E",
    smooth=False,
    marker_symbol="square",
    marker_size=10,
    # You can control where to place the data and the chart in the sheet
    # row_start=40,
    # column_start=3,
    # xlabel="Publication year",
    ylabel1="Publication count",
    ylabel2="World share",
    width=25,
    height=10,
    chart_position="bottom",
    rotation=0,
)

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

# NOTE: The baseline is detected and formatted differently. A baseline is a reference line with a constant value.
chart = LineChart()
chart.plot(
    df,
    ws,
    # Using the default openpyxl color
    openpyxl_color=True,
    section_heading="Using openpyxl default colors",
)

chart = LineChart()
chart.plot(
    df,
    ws,
    section_heading="Some customizations",
    title="Monthly sales",
    line_width=3,
    line_style="sysDash",
    # line_color="FF0000", Do not use this unless you have one line
    marker_symbol="circle",
    marker_size=8,
    # row_start=40,
    column_start=3,  # the column to start with
    xlabel="Month",
    ylabel="Sales",
    width=25,
    height=10,
    chart_position="bottom",
    rotation=30,  # x-axis major ticks
    # show_legend=False,  # You can disable legend
    legend_position="t",  # legend position: 'r', 't', 'l', 'b'
    # openpyxl_color=True,
    axis_bold_font=True,
)


chart = BarChart()
chart.plot(df, ws, width=30, section_heading="SQL2Excel colors")
chart.plot(df, ws, width=30, openpyxl_color=True, section_heading="Openpyxl colors")
chart.plot(df, ws, chart_type="bar")

chart = BarChart()
chart.plot(
    df,
    ws,
    section_heading="selecting specific columns - consecutive",
    title="Monthly sales",
    headings=[
        "Date",
        "Desktop sales",
        "Laptop sales",
        "Basline",
        "Tablet sales",
        "Phone sales",
    ],
    data_column_start=3,
    data_column_end=5,
    # data_columns=[2, 4, 5],
    # line_width=3,
    # line_style="sysDash",
    # line_color="FF0000",
    # smooth=False,
    # marker_symbol="circle",
    # marker_size=8,
    # row_start=40,
    column_start=3,
    xlabel="Month",
    ylabel="Sales",
    width=25,
    height=10,
    chart_position="bottom",
    rotation=90,
    # show_legend=False,
    legend_position="t",
)

chart.plot(
    df,
    ws,
    section_heading="selecting specific columns - non-consecutive",
    title="Monthly sales",
    headings=[
        "Date",
        "Desktop sales",
        "Laptop sales",
        "Basline",
        "Tablet sales",
        "Phone sales",
    ],
    data_columns=[2, 5],
    column_start=3,
    xlabel="Month",
    ylabel="Sales",
    width=25,
    height=10,
    chart_position="bottom",
    rotation=90,
    legend_position="t",
)

chart.plot(
    df,
    ws,
    section_heading="Customizing the axes",
    title="Monthly sales",
    title_font_name="Times New Roman",
    title_font_size=2400,
    title_font_color="green",
    title_font_bold=True,
    headings=[
        "Date",
        "Desktop sales",
        "Laptop sales",
        "Basline",
        "Tablet sales",
        "Phone sales",
    ],
    xlabel="Month",
    ylabel="Sales",
    axis_font_name="Times New Roman",
    axis_font_size=1800,  # 18pt
    axis_font_color="red",
    axis_bold_font=True,
    width=25,
    height=10,
    chart_position="bottom",
    rotation=90,
    # show_legend=False,
    legend_position="t",
)

# Radar Chart
total_sale_df = pd.DataFrame(
    {
        "Category": ["Desktop", "Laptop", "Tablet", "Phone"],
        "Total_Sale_First_Half": [
            round(df.iloc[:6][col].sum(), 0)
            for col in ["Desktop", "Laptop", "Tablet", "Phone"]
        ],
        "Total_Sale_Second_Half": [
            round(df.iloc[6:][col].sum(), 0)
            for col in ["Desktop", "Laptop", "Tablet", "Phone"]
        ],
    }
)

chart = RadarChart()
chart.plot(total_sale_df, ws, width=13, height=13, rotation=0)
chart.plot(total_sale_df, ws, chart_type="filled", width=13, height=13, rotation=0)

# Negative bars (Negative Sales!!)
df.iloc[:, -2] = -1 * df.iloc[:, -2]

chart = BarChart()
chart.plot(
    df,
    ws,
    section_heading="Monthly sales (xlabel position at the bottom because of negative bars)",
    xtick_label_position="low",
)

df = pd.DataFrame(
    {
        "pubyear": np.arange(2000, 2000 + 10),
        "pubcount": [100] * 10,
    }
)

chart = BarChart()
chart.plot(
    df,
    ws,
    vary_color=True,
    chart_position="bottom",
    width=25,
    height=10,
    section_heading="Color scheme",
    # row_start=60,
    column_start=3,
)

values = [1, 10, 100, 1000, 10000, 100000, 1000000]
n_rows = len(values)
df = pd.DataFrame(
    {
        "pubyear": list(range(2000, 2000 + n_rows)),
        # "pubcount": np.random.randint(1, 11, size=n_rows),
        "values": values,
    }
)

chart = LineChart()
chart.plot(df, ws, section_heading="Log scale (without log)", smooth=False)

chart.plot(df, ws, section_heading="Log scale (with log)", y_log_base=10, smooth=False)

rows = [
    [2, 40, 30],
    [3, 40, 25],
    [4, 50, 30],
    [5, 30, 10],
    [6, 25, 5],
    [7, 50, 10],
]

df = pd.DataFrame(data=rows, columns=["Number", "Batch 1", "Batch 2"])
df = df[["Number", "Batch 2", "Batch 1"]]
chart = AreaChart()
chart.plot(
    df,
    ws,
    section_heading="Area Chart",
    show_legend=True,
    shape_line_color="FFFFFF",
)

data = [
    [1, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55],
    [2, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65],
    [3, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75],
    [4, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85],
]

df = pd.DataFrame(data, columns=["X"] + [f"Series {i}" for i in range(1, 11)])
chart = ScatterChart()
chart.plot(
    df,
    ws,
    show_legend=True,
    nofill=True,
    marker_size=8,
    section_heading="Scatter plot without filling",
)
chart.plot(
    df,
    ws,
    show_legend=True,
    nofill=False,
    marker_size=7,
    section_heading="Scatter plot with filling",
)

# Color Scheme
df = pd.DataFrame(data=[[100 for _ in range(11)]], columns=list(range(11)))

chart = BarChart()

chart.plot(df, ws, width=20, show_legend=False)
chart.plot(df, ws, width=20, show_legend=False, openpyxl_color=True)


file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "df2report.xlsx"
)
wb.save(file_name)
