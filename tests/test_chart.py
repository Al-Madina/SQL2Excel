import matplotlib

matplotlib.use("Agg")  # Must be set before importing matplotlib.pyplot


import matplotlib.pyplot as plt
import openpyxl as xl
import pandas as pd
import pytest
from matplotlib.pyplot import bar
from sqlalchemy import all_

from sql2excel.chart import (
    AreaChart,
    BarChart,
    BarLineChart,
    BubbleChart,
    Chart,
    ImageChart,
    LineChart,
    PieChart,
    RadarChart,
    ScatterChart,
    SingleAxisBarLineChart,
    StackedBarChart,
)

from .setup import *
from .test_utils import *


def __assert_write_dataframe(df, worksheet, row_start, has_section_heading=False):
    offset = 1 if has_section_heading else 0
    assert worksheet.cell(row_start + offset, 1).value == "pubyear"
    assert worksheet.cell(row_start + offset, 2).value == "pubcount"
    assert worksheet.cell(row_start + offset + 1, 1).value == 2000
    assert worksheet.cell(row_start + len(df) + offset, 1).value == 2009


def test_write_dataframe_default(df, worksheet, chart, config):
    row_start = worksheet.max_row
    row_start = row_start + config.SEPARATOR + 1 if row_start > 1 else row_start
    chart.write_dataframe(df, worksheet)
    __assert_write_dataframe(df, worksheet, row_start)


def test_write_dataframe_with_section_heading(df, worksheet, chart, config):
    row_start = worksheet.max_row
    row_start = row_start + config.SEPARATOR + 1 if row_start > 1 else row_start
    chart.write_dataframe(df, worksheet, section_heading="Writing dataframe below")
    assert worksheet.cell(row_start, 1).value == "Writing dataframe below"
    __assert_write_dataframe(df, worksheet, row_start, has_section_heading=True)


def test_write_dataframe_with_dataframe_headings(df, worksheet, chart, config):
    row_start = worksheet.max_row
    row_start = row_start + config.SEPARATOR + 1 if row_start > 1 else row_start
    chart.write_dataframe(
        df, worksheet, headings=["publication year", "publication count", "percentage"]
    )
    assert worksheet.cell(row_start, 1).value == "publication year"
    assert worksheet.cell(row_start, 2).value == "publication count"
    assert worksheet.cell(row_start, 3).value == "percentage"
    assert worksheet.cell(row_start + 1, 1).value == 2000
    assert worksheet.cell(row_start + len(df), 1).value == 2009


def __assert_write_dataframe_side_by_side(
    dataframes, worksheet, row_start, data_data_separator, has_section_heading=False
):
    offset = 1 if has_section_heading else 0

    column_start = 1
    for df in dataframes:
        assert worksheet.cell(row_start + offset, column_start).value == "Col1"
        assert worksheet.cell(row_start + offset, column_start + 1).value == "Col2"
        assert (
            worksheet.cell(row_start + offset + 1, column_start).value == df.iloc[0, 0]
        )
        assert (
            worksheet.cell(row_start + offset + 1, column_start + 1).value
            == df.iloc[0, 1]
        )
        assert (
            worksheet.cell(row_start + offset + len(df), column_start).value
            == df.iloc[-1, 0]
        )
        assert (
            worksheet.cell(row_start + offset + len(df), column_start + 1).value
            == df.iloc[-1, 1]
        )
        column_start += len(df.columns) + data_data_separator


@pytest.mark.parametrize("num_dataframes", [2, 3, 4])
def test_write_dataframe_side_by_side_default(
    random_dataframes, worksheet, chart, config, num_dataframes
):
    # df1 = df
    # # df2 is the reverse of df1
    # df2 = df.iloc[::-1]
    dataframes = random_dataframes(num_dataframes)
    row_start = worksheet.max_row
    row_start = row_start + config.SEPARATOR + 1 if row_start > 1 else row_start
    chart.write_dataframes_side_by_side(dataframes, worksheet)
    __assert_write_dataframe_side_by_side(
        dataframes,
        worksheet,
        row_start,
        config.DATA_DATA_SEPARATOR,
        has_section_heading=False,
    )


@pytest.mark.parametrize("num_dataframes", [2, 3, 4])
def test_write_dataframe_side_by_side_with_section_heading(
    random_dataframes, clear_worksheet, chart, config, num_dataframes
):
    dataframes = random_dataframes(num_dataframes)
    worksheet = clear_worksheet
    row_start = worksheet.max_row
    row_start = row_start + config.SEPARATOR + 1 if row_start > 1 else row_start
    chart.write_dataframes_side_by_side(
        dataframes, worksheet, section_heading="Writing two dataframes below"
    )
    assert worksheet.cell(row_start, 1).value == "Writing two dataframes below"
    __assert_write_dataframe_side_by_side(
        dataframes,
        worksheet,
        row_start,
        config.DATA_DATA_SEPARATOR,
        has_section_heading=True,
    )


def test_write_dataframe_side_by_side_with_dataframe_heading(
    df, clear_worksheet, chart, config
):
    worksheet = clear_worksheet
    df1 = df
    # df2 is the reverse of df1
    df2 = df.iloc[::-1]
    row_start = worksheet.max_row
    row_start = row_start + config.SEPARATOR + 1 if row_start > 1 else row_start
    chart.write_dataframes_side_by_side(
        (df1, df2), worksheet, df_headings=["First DataFrame", "Second DataFrame"]
    )

    assert worksheet.cell(row_start, 1).value == "First DataFrame"
    assert (
        worksheet.cell(
            row_start, len(df1.columns) + config.DATA_DATA_SEPARATOR + 1
        ).value
        == "Second DataFrame"
    )

    # Assertions for df1
    assert worksheet.cell(row_start + 1, 1).value == "pubyear"
    assert worksheet.cell(row_start + 1, 2).value == "pubcount"
    assert worksheet.cell(row_start + 2, 1).value == 2000
    assert worksheet.cell(row_start + len(df) + 1, 1).value == 2009
    # Assertions for df2
    column_start = len(df1.columns) + config.DATA_DATA_SEPARATOR + 1
    assert worksheet.cell(row_start + 1, column_start).value == "pubyear"
    assert worksheet.cell(row_start + 1, column_start + 1).value == "pubcount"
    assert worksheet.cell(row_start + 2, column_start).value == 2009
    assert worksheet.cell(row_start + len(df) + 1, column_start).value == 2000


def test_write_dataframe_side_by_side_with_headings(df, worksheet, chart, config):
    df1 = df
    # df2 is the reverse of df1
    df2 = df.iloc[::-1]
    row_start = worksheet.max_row
    row_start = row_start + config.SEPARATOR + 1 if row_start > 1 else row_start
    headings = [[f"Col_{j}{i}" for i in range(1, 4)] for j in range(1, 3)]
    chart.write_dataframes_side_by_side((df1, df2), worksheet, headings=headings)
    # Assertions for df1
    assert worksheet.cell(row_start, 1).value == "Col_11"
    assert worksheet.cell(row_start, 2).value == "Col_12"
    assert worksheet.cell(row_start, 3).value == "Col_13"
    assert worksheet.cell(row_start + 1, 1).value == 2000
    assert worksheet.cell(row_start + len(df), 1).value == 2009
    # Assertions for df2
    column_start = len(df1.columns) + config.DATA_DATA_SEPARATOR + 1
    assert worksheet.cell(row_start, column_start).value == "Col_21"
    assert worksheet.cell(row_start, column_start + 1).value == "Col_22"
    assert worksheet.cell(row_start, column_start + 2).value == "Col_23"
    assert worksheet.cell(row_start + 1, column_start).value == 2009
    assert worksheet.cell(row_start + len(df), column_start).value == 2000


def test_create_line_chart(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    line_chart.plot(df, worksheet)
    assert len(worksheet._charts) == 1
    ws_chart = worksheet._charts[0]
    column_letter = excel_helper.get_column_letter(
        len(df.columns) + config.DATA_CHART_SEPARATOR + 1
    )
    assert ws_chart.anchor == column_letter + str(1)


def test_line_chart_series_value(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    line_chart.plot(df, worksheet)
    ws_chart = worksheet._charts[0]

    series_values = get_series_values(ws_chart.series[0], worksheet)
    assert df.iloc[:, 1].tolist() == series_values


@pytest.mark.parametrize("chart_class", [LineChart, BarChart, PieChart, RadarChart])
@pytest.mark.parametrize(
    "column_range",
    [
        (1, 2),
        (1, 3),
        (2, 3),
        (2, 4),
        (2, 5),
        (1, 1),
        (2, 2),
        (3, 3),
    ],
)
@pytest.mark.parametrize("column_start", [None, 1, 2, 3, 4, 5])
def test_specify_data_columns_start_and_end(
    long_wide_df,
    create_worksheet,
    config,
    excel_helper,
    chart_class,
    column_range,
    column_start,
):
    """Testing ...
    NOTE: When specifying `column_start` the reference to series in the sheet is offset by
    `column_start`."""
    df = long_wide_df
    worksheet = create_worksheet
    chart = chart_class(config=config, excel_helper=excel_helper)

    data_column_start, data_column_end = column_range
    chart.plot(
        df,
        worksheet,
        data_column_start=data_column_start,
        data_column_end=data_column_end,
        column_start=column_start,
    )

    ws_chart = worksheet._charts[0]

    # When specifying column start and end, the columns are contiguous
    for i, series in enumerate(ws_chart.series):
        series_values = get_series_values(series, worksheet)
        expected_values = df.iloc[:, data_column_start - 1 + i].tolist()
        assert series_values == expected_values


@pytest.mark.parametrize("chart_class", [LineChart, BarChart, PieChart, RadarChart])
@pytest.mark.parametrize("data_column_start", list(range(1, 6)))
@pytest.mark.parametrize("column_start", [None, 1, 2, 3, 4, 5])
def test_specify_data_columns_start_only(
    long_wide_df,
    create_worksheet,
    config,
    excel_helper,
    chart_class,
    data_column_start,
    column_start,
):
    df = long_wide_df
    worksheet = create_worksheet
    chart = chart_class(config=config, excel_helper=excel_helper)

    chart.plot(
        df, worksheet, data_column_start=data_column_start, column_start=column_start
    )

    ws_chart = worksheet._charts[0]

    # When specifying column start and end, the columns are contiguous
    for i, series in enumerate(ws_chart.series):
        series_values = get_series_values(series, worksheet)
        expected_values = df.iloc[:, data_column_start - 1 + i].tolist()
        assert series_values == expected_values


@pytest.mark.parametrize(
    "chart_class",
    [LineChart, BarChart, PieChart, RadarChart],
)
@pytest.mark.parametrize(
    "data_columns",
    [[1], [2], [3], [1, 2], [1, 3], [1, 5], [2, 3], [1, 2, 3], [2, 3, 5]],
)
@pytest.mark.parametrize("column_start", [None, 1, 2, 3, 4, 5])
def test_specify_data_columns(
    long_wide_df,
    create_worksheet,
    config,
    excel_helper,
    chart_class,
    data_columns,
    column_start,
):
    df = long_wide_df
    worksheet = create_worksheet
    chart = chart_class(config=config, excel_helper=excel_helper)

    # data_columns = [2, 3, 5]
    chart.plot(df, worksheet, data_columns=data_columns, column_start=column_start)

    ws_chart = worksheet._charts[0]
    for i, series in enumerate(ws_chart.series):
        series_values = get_series_values(series, worksheet)
        expected_values = df.iloc[:, data_columns[i] - 1].tolist()
        assert series_values == expected_values


def test_single_axis_bar_line_chart_default_columns(
    long_wide_df, create_worksheet, config, excel_helper
):
    df = long_wide_df
    worksheet = create_worksheet
    chart = SingleAxisBarLineChart(
        config=config,
        excel_helper=excel_helper,
    )

    chart.plot(df, worksheet)

    assert len(worksheet._charts) == 1
    ws_chart = worksheet._charts[0]

    # In composite charts, series are spread across sub-charts.
    all_series = []
    print(len(all_series))
    print(len(list(ws_chart._charts)))
    for subchart in ws_chart._charts:
        all_series.extend(subchart.series)

    assert len(all_series) == 5  # 1 bar series + 4 line series

    # Check values for bar (Series 0 -> Col 2)
    assert get_series_values(all_series[0], worksheet) == df.iloc[:, 1].tolist()
    # Check values for lines (Series 1-4 -> Col 3-6)
    for i in range(1, 5):
        assert get_series_values(all_series[i], worksheet) == df.iloc[:, i + 1].tolist()


def test_single_axis_bar_line_chart_explicit_columns(
    long_wide_df, create_worksheet, config, excel_helper
):
    df = long_wide_df
    worksheet = create_worksheet
    chart = SingleAxisBarLineChart(config=config, excel_helper=excel_helper)

    # 2 bars, 3 lines
    chart.plot(df, worksheet, bar_columns=[2, 5], line_columns=[3, 4, 6])

    ws_chart = worksheet._charts[0]

    all_series = []
    for subchart in ws_chart._charts:
        all_series.extend(subchart.series)  # Line series
    assert len(all_series) == 5

    # Series order: bars first, then lines
    assert get_series_values(all_series[0], worksheet) == df.iloc[:, 1].tolist()
    assert get_series_values(all_series[1], worksheet) == df.iloc[:, 4].tolist()
    assert get_series_values(all_series[2], worksheet) == df.iloc[:, 2].tolist()
    assert get_series_values(all_series[3], worksheet) == df.iloc[:, 3].tolist()
    assert get_series_values(all_series[4], worksheet) == df.iloc[:, 5].tolist()


def test_single_axis_bar_line_chart_styling(
    long_wide_df, create_worksheet, config, excel_helper
):
    df = long_wide_df
    worksheet = create_worksheet
    chart = SingleAxisBarLineChart(config=config, excel_helper=excel_helper)

    chart.plot(
        df,
        worksheet,
        line_width=2.5,
        line_style="dash",
        marker_symbol="diamond",
        marker_size=10,
    )

    ws_chart = worksheet._charts[0]

    all_series = list(ws_chart.series)  # Bar series
    for subchart in ws_chart._charts:  # Assuming only one subchart (the line chart)
        all_series.extend(subchart.series)  # Line series

    # Line series start at index 2 (since there are 2 bar columns: [2, 5])
    for i in range(2, 5):  # Indices 2, 3, 4 correspond to the 3 line series
        series = all_series[i]
        assert series.graphicalProperties.line.width == 2.5 * 12700
        assert series.graphicalProperties.line.dashStyle == "dash"
        assert series.marker.symbol == "diamond"
        assert series.marker.size == 10


def test_base_chart_add_image_not_implemented(chart):
    with pytest.raises(NotImplementedError):
        chart.add_image("fake.png", None)


def test_base_chart_plot_not_implemented(chart):
    with pytest.raises(NotImplementedError):
        chart.plot(None, None)


def test_line_chart_from_rows_with_data_column_start_raises(
    df, create_worksheet, config, excel_helper
):
    worksheet = create_worksheet
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    with pytest.raises(NotImplementedError):
        line_chart.plot(df, worksheet, from_rows=True, data_column_start=1)


def test_line_chart_from_rows_with_data_columns_raises(
    df, create_worksheet, config, excel_helper
):
    worksheet = create_worksheet
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    with pytest.raises(NotImplementedError):
        line_chart.plot(df, worksheet, from_rows=True, data_columns=[1, 2])


def test_line_chart_from_rows_warns_single_unique_value(
    create_worksheet, config, excel_helper
):
    worksheet = create_worksheet
    single_unique_df = pd.DataFrame(
        {
            "category": ["A", "B", "C"],
            "value": [10, 20, 30],
            "constant": [42, 42, 42],
        }
    )
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    with pytest.warns(UserWarning):
        line_chart.plot(single_unique_df, worksheet, from_rows=True)


def test_create_bar_chart(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    bar_chart = BarChart(config=config, excel_helper=excel_helper)
    bar_chart.plot(df, worksheet)
    assert len(worksheet._charts) == 1
    ws_chart = worksheet._charts[0]
    assert isinstance(ws_chart, xl.chart.BarChart)


def test_bar_chart_series_values(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    bar_chart = BarChart(config=config, excel_helper=excel_helper)
    bar_chart.plot(df, worksheet)
    ws_chart = worksheet._charts[0]
    series_values = get_series_values(ws_chart.series[0], worksheet)
    assert df.iloc[:, 1].tolist() == series_values


def test_bar_chart_chart_type_bar(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    bar_chart = BarChart(config=config, excel_helper=excel_helper)
    bar_chart.plot(df, worksheet, chart_type="bar")
    ws_chart = worksheet._charts[0]
    assert ws_chart.type == "bar"


def test_bar_chart_chart_type_col(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    bar_chart = BarChart(config=config, excel_helper=excel_helper)
    bar_chart.plot(df, worksheet, chart_type="col")
    ws_chart = worksheet._charts[0]
    assert ws_chart.type == "col"


def test_bar_chart_xtick_label_position(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    bar_chart = BarChart(config=config, excel_helper=excel_helper)
    bar_chart.plot(df, worksheet, xtick_label_position="low")
    ws_chart = worksheet._charts[0]
    assert ws_chart.x_axis.tickLblPos == "low"
    assert ws_chart.x_axis.tickLblSkip == 2


def test_bar_chart_ytick_label_position(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    bar_chart = BarChart(config=config, excel_helper=excel_helper)
    bar_chart.plot(df, worksheet, ytick_label_position="low")
    ws_chart = worksheet._charts[0]
    assert ws_chart.y_axis.tickLblPos == "low"
    assert ws_chart.y_axis.tickLblSkip == 2


def test_create_stacked_bar_chart(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = StackedBarChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    assert len(worksheet._charts) == 1
    ws_chart = worksheet._charts[0]
    assert isinstance(ws_chart, xl.chart.BarChart)


def test_stacked_bar_chart_default_grouping(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = StackedBarChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    ws_chart = worksheet._charts[0]
    assert ws_chart.grouping == "percentStacked"
    assert ws_chart.overlap == 100
    assert ws_chart.style == 10


def test_stacked_bar_chart_stacked_grouping(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = StackedBarChart(config=config, excel_helper=excel_helper)
    chart.plot(
        df, worksheet, chart_grouping="stacked", chart_overlap=50, chart_style=15
    )
    ws_chart = worksheet._charts[0]
    assert ws_chart.grouping == "stacked"
    assert ws_chart.overlap == 50
    assert ws_chart.style == 15


def test_stacked_bar_chart_series_values(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = StackedBarChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    ws_chart = worksheet._charts[0]
    series_values = get_series_values(ws_chart.series[0], worksheet)
    assert df.iloc[:, 1].tolist() == series_values


def test_create_pie_chart(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = PieChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    assert len(worksheet._charts) == 1
    ws_chart = worksheet._charts[0]
    assert isinstance(ws_chart, xl.chart.PieChart)


def test_pie_chart_series_values(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = PieChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    ws_chart = worksheet._charts[0]
    series_values = get_series_values(ws_chart.series[0], worksheet)
    assert df.iloc[:, 1].tolist() == series_values


def test_pie_chart_default_data_labels(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = PieChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    ws_chart = worksheet._charts[0]
    data_labels = ws_chart.dataLabels
    assert data_labels.showCatName is True
    assert data_labels.showPercent is True
    assert data_labels.showVal is False
    assert data_labels.showSerName is False
    assert data_labels.showLegendKey is False


def test_pie_chart_custom_data_labels(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = PieChart(config=config, excel_helper=excel_helper)
    chart.plot(
        df,
        worksheet,
        show_percentage=False,
        show_category=False,
        show_values=True,
        show_series_name=True,
        show_legend_key=True,
    )
    ws_chart = worksheet._charts[0]
    data_labels = ws_chart.dataLabels
    assert data_labels.showCatName is False
    assert data_labels.showPercent is False
    assert data_labels.showVal is True
    assert data_labels.showSerName is True
    assert data_labels.showLegendKey is True


def test_create_area_chart(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = AreaChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    assert len(worksheet._charts) == 1
    ws_chart = worksheet._charts[0]
    assert isinstance(ws_chart, xl.chart.AreaChart)


def test_area_chart_series_values(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = AreaChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    ws_chart = worksheet._charts[0]
    series_values = get_series_values(ws_chart.series[0], worksheet)
    assert df.iloc[:, 1].tolist() == series_values


def test_area_chart_custom_style(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = AreaChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet, chart_style=20)
    ws_chart = worksheet._charts[0]
    assert ws_chart.style == 20


def test_area_chart_default_style(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = AreaChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    ws_chart = worksheet._charts[0]
    assert ws_chart.style == 13


def test_create_scatter_chart(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = ScatterChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    assert len(worksheet._charts) == 1
    ws_chart = worksheet._charts[0]
    assert isinstance(ws_chart, xl.chart.ScatterChart)


def test_scatter_chart_default_series(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = ScatterChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    ws_chart = worksheet._charts[0]
    assert len(ws_chart.series) == 2


def test_scatter_chart_data_column_start(
    long_wide_df, create_worksheet, config, excel_helper
):
    df = long_wide_df
    worksheet = create_worksheet
    chart = ScatterChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet, data_column_start=2, data_column_end=3)
    ws_chart = worksheet._charts[0]
    assert len(ws_chart.series) == 2


def test_scatter_chart_data_columns(
    long_wide_df, create_worksheet, config, excel_helper
):
    df = long_wide_df
    worksheet = create_worksheet
    chart = ScatterChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet, data_columns=[2, 4])
    ws_chart = worksheet._charts[0]
    assert len(ws_chart.series) == 2


def test_scatter_chart_from_rows_raises(
    long_wide_df, create_worksheet, config, excel_helper
):
    df = long_wide_df
    worksheet = create_worksheet
    chart = ScatterChart(config=config, excel_helper=excel_helper)
    with pytest.raises(NotImplementedError):
        chart.plot(df, worksheet, from_rows=True)


def test_create_bubble_chart(long_wide_df, create_worksheet, config, excel_helper):
    df = long_wide_df
    worksheet = create_worksheet
    chart = BubbleChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    assert len(worksheet._charts) == 1
    ws_chart = worksheet._charts[0]
    assert isinstance(ws_chart, xl.chart.BubbleChart)
    assert len(ws_chart.series) == len(df)


def test_bubble_chart_series_titles(
    long_wide_df, create_worksheet, config, excel_helper
):
    df = long_wide_df
    worksheet = create_worksheet
    chart = BubbleChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    ws_chart = worksheet._charts[0]
    for idx, series in enumerate(ws_chart.series):
        assert int(series.title.v) == df.iloc[idx, 0]


def test_radar_chart_default_type(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = RadarChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    ws_chart = worksheet._charts[0]
    assert ws_chart.type == "standard"


def test_radar_chart_chart_type_marker(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = RadarChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet, chart_type="marker")
    ws_chart = worksheet._charts[0]
    assert ws_chart.type == "marker"


def test_radar_chart_chart_type_filled(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = RadarChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet, chart_type="filled")
    ws_chart = worksheet._charts[0]
    assert ws_chart.type == "filled"


def test_radar_chart_y_axis_scaling(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    chart = RadarChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet, radar_unit=0.5)
    ws_chart = worksheet._charts[0]
    assert ws_chart.y_axis.majorUnit == 0.5


def test_line_chart_title(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    line_chart.plot(df, worksheet, title="Publication Trends")
    ws_chart = worksheet._charts[0]
    assert ws_chart.title.tx.rich.p[0].r[0].t == "Publication Trends"


def test_line_chart_xlabel_ylabel(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    line_chart.plot(df, worksheet, xlabel="Year", ylabel="Count")
    ws_chart = worksheet._charts[0]
    assert ws_chart.x_axis.title.tx.rich.p[0].r[0].t == "Year"
    assert ws_chart.y_axis.title.tx.rich.p[0].r[0].t == "Count"


def test_line_chart_no_legend(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    line_chart.plot(df, worksheet, show_legend=False)
    ws_chart = worksheet._charts[0]
    assert ws_chart.legend is None


def test_line_chart_position_bottom(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    line_chart.plot(df, worksheet, chart_position="bottom")
    ws_chart = worksheet._charts[0]
    expected_col = xl.utils.cell.get_column_letter(line_chart.min_col - 1)
    expected_row = line_chart.max_row + config.DATA_CHART_SEPARATOR + 1
    assert ws_chart.anchor == expected_col + str(expected_row)


def test_line_chart_width_height(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    line_chart.plot(df, worksheet, width=15, height=10)
    ws_chart = worksheet._charts[0]
    assert ws_chart.width == 15
    assert ws_chart.height == 10


def test_line_chart_chart_style(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    line_chart.plot(df, worksheet, chart_style=20)
    ws_chart = worksheet._charts[0]
    assert ws_chart.style == 20


def test_image_chart_matplotlib_figure(create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    image_chart = ImageChart(config=config, excel_helper=excel_helper)
    fig, ax = plt.subplots()
    ax.plot([1, 2, 3], [1, 2, 3])
    image_chart.add_image(fig, worksheet, chart_position="bottom")
    assert len(worksheet._images) == 1


def test_image_chart_with_df(df, create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    image_chart = ImageChart(config=config, excel_helper=excel_helper)
    fig, ax = plt.subplots()
    ax.plot([1, 2, 3], [1, 2, 3])
    image_chart.add_image(fig, worksheet, df=df, chart_position="bottom")
    assert worksheet.cell(1, 1).value == "pubyear"
    assert worksheet.cell(1, 2).value == "pubcount"
    assert len(worksheet._images) == 1


def test_image_chart_unsupported_type_warns(create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    image_chart = ImageChart(config=config, excel_helper=excel_helper)
    with pytest.warns(UserWarning, match="Unsupported image input type"):
        image_chart.add_image(123, worksheet, chart_position="bottom")


def test_image_chart_unknown_position_warns(create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    image_chart = ImageChart(config=config, excel_helper=excel_helper)
    fig, ax = plt.subplots()
    ax.plot([1, 2, 3], [1, 2, 3])
    with pytest.warns(UserWarning, match="Unknown chart position"):
        image_chart.add_image(fig, worksheet, chart_position="unknown")


def test_image_chart_with_section_heading(create_worksheet, config, excel_helper):
    worksheet = create_worksheet
    image_chart = ImageChart(config=config, excel_helper=excel_helper)
    fig, ax = plt.subplots()
    ax.plot([1, 2, 3], [1, 2, 3])
    image_chart.add_image(
        fig, worksheet, chart_position="bottom", section_heading="My Image"
    )
    assert worksheet.cell(1, 1).value == "My Image"
    assert len(worksheet._images) == 1


def test_create_bar_line_chart(long_wide_df, create_worksheet, config, excel_helper):
    df = long_wide_df
    worksheet = create_worksheet
    chart = BarLineChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    assert len(worksheet._charts) == 1


def test_bar_line_chart_series(long_wide_df, create_worksheet, config, excel_helper):
    df = long_wide_df
    worksheet = create_worksheet
    chart = BarLineChart(config=config, excel_helper=excel_helper)
    chart.plot(df, worksheet)
    assert len(chart.chart1.series) + len(chart.chart2.series) > 0
