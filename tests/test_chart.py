# TODO complete the tests
# TODO Test chart position (right, bottom, unknown)
# TODO Test placing dataframe and chart at specific (x, y)
# TODO Test reference column

import pytest
from setup import *

from sql2excel.chart import (
    BarChart,
    Chart,
    LineChart,
    PieChart,
    RadarChart,
    # BarLineChart,
)
from tests.test_utils import *


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
