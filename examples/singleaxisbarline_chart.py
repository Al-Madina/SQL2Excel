# An example of different single axis barline charts
import os

import numpy as np
import openpyxl as xl
import pandas as pd

from sql2excel.chart import SingleAxisBarLineChart

wb = xl.Workbook()
ws = wb.active

n_rows = 11

df = pd.DataFrame(
    {
        "pubyear": np.arange(2000, 2000 + n_rows),
        "pubcount": [100 * (i / 2 + 1) for i in range(n_rows)],
        "share": [100 + (i * np.random.rand()) for i in range(n_rows)],
    }
)

# Add more columns for testing different bar/line combinations
df["share2"] = 1.5 * df["share"]
df["pubcount2"] = [50 * (i / 2 + 1) for i in range(n_rows)]
df["share3"] = 1.9 * df["share"]

# Basic plot with default behavior:
# pubcount will be bar, share, share2, pubcount2, share3 will be lines
chart = SingleAxisBarLineChart()
chart.plot(
    df.drop(columns=["pubcount2"]),
    ws,
    section_heading="Single Axis Bar Line Chart: Default behavior",
    title="Publication count and world share (Default)",
    ylabel="Value",
)

# Custom colors
chart = SingleAxisBarLineChart()
chart.plot(
    df.drop(columns=["pubcount2"]),
    ws,
    section_heading="Single Axis Bar Line Chart: Custom colors (SQL2Excel palette)",
    title="Publication count and world share (Custom Colors)",
    ylabel="Value",
    custom_colors=True,
)

# One bar, one line (explicitly defined)
chart = SingleAxisBarLineChart()
chart.plot(
    df,
    ws,
    section_heading="Single Axis Bar Line Chart: One bar (pubcount), one line (share) - User-specified columns",
    title="Publication count and world share (Explicit One Bar, One Line)",
    bar_columns=[2],  # pubcount
    line_columns=[3],  # share
    ylabel="Value",
)

# One bar, one line (inferred)
chart = SingleAxisBarLineChart()
chart.plot(
    df[["pubyear", "pubcount", "share"]],
    ws,
    section_heading="Single Axis Bar Line Chart: One bar (pubcount), one line (share) - Inferred",
    title="Publication count and world share (Explicit One Bar, One Line)",
    line_style="sysDash",
    ylabel="Value",
)
# One bar, multiple lines
chart = SingleAxisBarLineChart()
chart.plot(
    df,
    ws,
    section_heading="Single Axis Bar Line Chart: One bar (pubcount), multiple lines (share, share2, share3)",
    title="Publication count and world share (One Bar, Multiple Lines) - User-specified columns",
    bar_columns=[2],  # pubcount
    line_columns=[3, 4, 6],  # share, share2, share3
    ylabel="Value",
)

# Multiple bars, one line
chart = SingleAxisBarLineChart()
chart.plot(
    df,
    ws,
    section_heading="Single Axis Bar Line Chart: Multiple bars (pubcount, pubcount2), one line (share) - User-specified columns",
    title="Publication count and world share (Multiple Bars, One Line)",
    bar_columns=[2, 5],  # pubcount, pubcount2
    line_columns=[3],  # share
    ylabel="Value",
)

# Multiple bars, multiple lines
chart = SingleAxisBarLineChart()
chart.plot(
    df,
    ws,
    section_heading="Single Axis Bar Line Chart: Multiple bars (pubcount, pubcount2), multiple lines (share, share2, share3)",
    title="Publication count and world share (Multiple Bars, Multiple Lines)",
    bar_columns=[2, 5],  # pubcount, pubcount2
    line_columns=[3, 4, 6],  # share, share2, share3
    ylabel="Value",
)

# Inferred line_columns (only bar_columns provided)
chart = SingleAxisBarLineChart()
chart.plot(
    df.drop(columns=["pubcount2"]),
    ws,
    section_heading="Single Axis Bar Line Chart: Inferred line_columns (pubcount as bar)",
    title="Publication count and world share (Inferred Lines)",
    bar_columns=[2],  # pubcount
    ylabel="Value",
)

# Inferred bar_columns (only line_columns provided)
chart = SingleAxisBarLineChart()
chart.plot(
    df,
    ws,
    section_heading="Single Axis Bar Line Chart: Inferred bar_columns (share as line)",
    title="Publication count and world share (Inferred Bars)",
    line_columns=[3],  # share
    ylabel="Value",
)

# Specific line_color with custom_colors=False
chart = SingleAxisBarLineChart()
chart.plot(
    df,
    ws,
    section_heading="Single Axis Bar Line Chart: Specific line color (custom_colors=False)",
    title="Publication count and world share (Specific Line Color)",
    bar_columns=[2],
    line_columns=[3],
    # line_color="107C10",  # Dark red
    custom_colors=False,
    ylabel="Value",
)

# Extensive customization
chart = SingleAxisBarLineChart()
chart.plot(
    df,
    ws,
    section_heading="Single Axis Bar Line Chart: With extensive customization",
    title="Publication Trends and Shares (Customized)",
    headings=["Year", "Pub Count", "Share 1", "Share 2", "Pub Count 2", "Share 3"],
    bar_columns=[2, 5],  # pubcount, pubcount2
    line_columns=[3, 4, 6],  # share, share2, share3
    line_width=2,
    line_style="sysDash",
    smooth=True,
    marker_symbol="diamond",
    marker_size=7,
    column_start=3,
    ylabel="Combined Value",
    ylim=(0, 1200),
    width=25,
    height=12,
    chart_position="bottom",
    rotation=45,
    axis_bold_font=True,
    custom_colors=True,  # Use SQL2Excel custom colors for distinct lines/bars
)

file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "singleaxisbarline_chart.xlsx"
)
wb.save(file_name)
