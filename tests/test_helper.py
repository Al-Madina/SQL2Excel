import pandas as pd
import pytest
from openpyxl.chart import LineChart as XlLineChart

from sql2excel.chart import LineChart

from .setup import *


def test_is_sheet_empty(clear_worksheet, excel_helper):
    worksheet = clear_worksheet
    assert excel_helper.is_sheet_empty(worksheet)


def test_is_sheet_empty_with_non_empty_sheet(add_content_to_sheet, excel_helper):
    worksheet = add_content_to_sheet
    assert not excel_helper.is_sheet_empty(worksheet)


def test_row_start_with_empty_sheet(clear_worksheet, excel_helper):
    worksheet = clear_worksheet
    assert excel_helper.is_sheet_empty(worksheet)
    row_start = excel_helper.get_row_start(worksheet)
    assert row_start == 1


def test_row_start_with_non_empty_sheet(add_content_to_sheet, config, excel_helper):
    worksheet = add_content_to_sheet
    current_max_row = worksheet.max_row
    assert (
        excel_helper.get_row_start(worksheet) == current_max_row + config.SEPARATOR + 1
    )


@pytest.mark.parametrize(
    "col, expected",
    [
        (None, "A"),
        ("A", "A"),
        ("Z", "Z"),
        ("AA", "AA"),
        (1, "A"),
        (26, "Z"),
        (27, "AA"),
        (16384, "XFD"),
        # NOTE a decision was made to prevent raising exceptions resulted from excel-related
        # code. a UserWarning is giving instead and a default config/value is used.
        # ("XFE", ValueError),
        # (16385, ValueError),
        # (-1, ValueError),
        # ("$", ValueError),
        # ("1", ValueError),
    ],
)
def test_get_column_letter(excel_helper, col, expected):
    if isinstance(expected, type) and issubclass(expected, Exception):
        with pytest.raises(expected):
            excel_helper.get_column_letter(col)
    else:
        result = excel_helper.get_column_letter(col)
        assert result == expected


@pytest.mark.parametrize(
    "col, expected",
    [
        (0, "A"),
        (-1, "A"),
        (16385, "A"),
    ],
)
def test_get_column_letter_invalid_int_warns(excel_helper, col, expected):
    with pytest.warns(UserWarning, match="Invalid column"):
        result = excel_helper.get_column_letter(col)
    assert result == expected


def test_get_column_letter_invalid_str_warns(excel_helper):
    with pytest.warns(UserWarning, match="Invalid column"):
        result = excel_helper.get_column_letter("XFE")
    assert result == "A"


def test_get_starting_position_empty_sheet(clear_worksheet, excel_helper):
    worksheet = clear_worksheet
    row, col = excel_helper.get_starting_position(worksheet)
    assert row == 1
    assert col == 1


def test_get_starting_position_non_empty_sheet(
    add_content_to_sheet, config, excel_helper
):
    worksheet = add_content_to_sheet
    current_max_row = worksheet.max_row
    row, col = excel_helper.get_starting_position(worksheet)
    assert row == current_max_row + config.SEPARATOR + 1
    assert col == 1


def test_get_starting_position_with_row_start(add_content_to_sheet, excel_helper):
    worksheet = add_content_to_sheet
    row, col = excel_helper.get_starting_position(worksheet, row_start=5)
    assert row == 5
    assert col == 1


def test_get_starting_position_with_column_start(add_content_to_sheet, excel_helper):
    worksheet = add_content_to_sheet
    row, col = excel_helper.get_starting_position(worksheet, column_start="B")
    assert col == 2


def test_get_starting_position_with_both(add_content_to_sheet, excel_helper):
    worksheet = add_content_to_sheet
    row, col = excel_helper.get_starting_position(
        worksheet, row_start=3, column_start="C"
    )
    assert row == 3
    assert col == 3


def test_get_starting_position_invalid_row_start_warns(
    add_content_to_sheet, excel_helper
):
    worksheet = add_content_to_sheet
    current_max_row = worksheet.max_row
    with pytest.warns(UserWarning, match="Row start must be an integer"):
        row, col = excel_helper.get_starting_position(worksheet, row_start=-1)
    assert row == current_max_row + excel_helper.config.SEPARATOR + 1


def test_get_starting_position_invalid_row_type_warns(
    add_content_to_sheet, excel_helper
):
    worksheet = add_content_to_sheet
    current_max_row = worksheet.max_row
    with pytest.warns(UserWarning, match="Row start must be an integer"):
        row, col = excel_helper.get_starting_position(worksheet, row_start="invalid")
    assert row == current_max_row + excel_helper.config.SEPARATOR + 1


def test_get_starting_position_with_int_column(excel_helper, clear_worksheet):
    worksheet = clear_worksheet
    row, col = excel_helper.get_starting_position(worksheet, column_start=3)
    assert col == 3


def test_insert_rows_for_chart_height_basic(clear_worksheet, excel_helper):
    worksheet = clear_worksheet
    excel_helper.insert_rows_for_chart_height(10, worksheet)
    assert worksheet.max_row > 0


def test_insert_rows_for_chart_height_with_df(create_worksheet, excel_helper, df):
    worksheet = create_worksheet
    n_rows = int(excel_helper.config.CHART_HEIGHT_SCALE * 10) + 1
    n_rows_df = df.shape[0]
    expected_appended = max(n_rows - n_rows_df, 0)
    excel_helper.insert_rows_for_chart_height(10, worksheet, df=df)
    assert worksheet.max_row == expected_appended


def test_insert_rows_for_chart_height_with_scale(clear_worksheet, excel_helper):
    worksheet = clear_worksheet
    excel_helper.insert_rows_for_chart_height(5, worksheet, scale=1.0)
    assert worksheet.max_row == int(1.0 * 5) + 1


def test_insert_rows_for_chart_height_df_larger_than_chart(
    clear_worksheet, excel_helper
):
    worksheet = clear_worksheet
    import pandas as pd

    big_df = pd.DataFrame({"a": range(100)})
    rows_before = worksheet.max_row
    excel_helper.insert_rows_for_chart_height(5, worksheet, df=big_df)
    assert worksheet.max_row == rows_before


def _make_series_for_helper(excel_helper, config):
    """Helper to create a line chart with a single series for testing graphical properties."""
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    ws = xl.Workbook().create_sheet()
    line_chart.plot(
        pd.DataFrame({"x": [1, 2], "y": [3, 4]}),
        ws,
        title="T",
    )
    return ws._charts[0].series[0]


def test_set_line_graphical_properties_width(config, excel_helper):
    series = _make_series_for_helper(excel_helper, config)
    excel_helper.set_line_graphical_properties(series, width=2.0)
    assert series.graphicalProperties.line.width == 2.0 * 12700


def test_set_line_graphical_properties_style(config, excel_helper):
    series = _make_series_for_helper(excel_helper, config)
    excel_helper.set_line_graphical_properties(series, style="dash")
    assert series.graphicalProperties.line.dashStyle == "dash"


def test_set_line_graphical_properties_color(config, excel_helper):
    series = _make_series_for_helper(excel_helper, config)
    excel_helper.set_line_graphical_properties(series, color="FF0000")
    assert series.graphicalProperties.line.solidFill is not None


def test_set_line_graphical_properties_smooth(config, excel_helper):
    series = _make_series_for_helper(excel_helper, config)
    excel_helper.set_line_graphical_properties(series, smooth=True)
    assert series.smooth is True


def test_set_line_graphical_properties_smooth_false(config, excel_helper):
    series = _make_series_for_helper(excel_helper, config)
    excel_helper.set_line_graphical_properties(series, smooth=False)
    assert series.smooth is False


def test_set_line_graphical_properties_nofill(config, excel_helper):
    series = _make_series_for_helper(excel_helper, config)
    excel_helper.set_line_graphical_properties(series, nofill=True)
    assert series.graphicalProperties.line.noFill is True


def test_set_marker_graphical_properties_symbol(config, excel_helper):
    series = _make_series_for_helper(excel_helper, config)
    excel_helper.set_marker_graphical_properties(series, symbol="circle")
    assert series.marker.symbol == "circle"
    assert series.marker.size == config.MARKER_SIZE


def test_set_marker_graphical_properties_symbol_and_size(config, excel_helper):
    series = _make_series_for_helper(excel_helper, config)
    excel_helper.set_marker_graphical_properties(series, symbol="diamond", size=10)
    assert series.marker.symbol == "diamond"
    assert series.marker.size == 10


def test_set_marker_graphical_properties_color(config, excel_helper):
    series = _make_series_for_helper(excel_helper, config)
    excel_helper.set_marker_graphical_properties(
        series, symbol="circle", color="FF0000"
    )
    assert series.marker.graphicalProperties.solidFill is not None


def test_fill_with_color(config, excel_helper):
    series = _make_series_for_helper(excel_helper, config)
    excel_helper.fill(series, color="FF0000")
    assert series.graphicalProperties.solidFill is not None
    assert series.graphicalProperties.line.solidFill is not None


def test_fill_with_border_line_color(config, excel_helper):
    series = _make_series_for_helper(excel_helper, config)
    excel_helper.fill(series, color="FF0000", border_line_color="0000FF")
    assert series.graphicalProperties.solidFill is not None
    assert series.graphicalProperties.line.solidFill is not None


def test_fill_data_point(config, excel_helper):
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    ws = xl.Workbook().create_sheet()
    line_chart.plot(
        pd.DataFrame({"x": [1, 2, 3], "y": [4, 5, 6]}),
        ws,
        title="T",
    )
    series = ws._charts[0].series[0]
    series_length = 3
    excel_helper.fill_data_point(series, series_length)
    assert len(series.data_points) == series_length


def test_fill_data_point_recycles_colors(config, excel_helper):
    line_chart = LineChart(config=config, excel_helper=excel_helper)
    ws = xl.Workbook().create_sheet()
    n_points = len(config.PRIMARY_COLORS) + 3
    line_chart.plot(
        pd.DataFrame({"x": list(range(n_points)), "y": list(range(n_points))}),
        ws,
        title="T",
    )
    series = ws._charts[0].series[0]
    excel_helper.fill_data_point(series, n_points)
    assert len(series.data_points) == n_points


def test_reference_column_exists_true(df, excel_helper):
    import pandas as pd

    test_df = pd.DataFrame({"a": [1, 2], "b": ["same", "same"]})
    assert excel_helper.reference_column_exists(test_df) is True


def test_reference_column_exists_false(df, excel_helper):
    import pandas as pd

    test_df = pd.DataFrame({"a": [1, 2], "b": ["diff1", "diff2"]})
    assert excel_helper.reference_column_exists(test_df) is False


def test_set_section_heading_font_defaults(clear_worksheet, config, excel_helper):
    worksheet = clear_worksheet
    worksheet.append(["Test"])
    cell = worksheet.cell(1, 1)
    excel_helper.set_section_heading_font(cell)
    assert cell.font.name == config.SECTION_HEADING_FONT_NAME
    assert cell.font.bold == config.SECTION_HEADING_BOLD
    assert cell.font.size == config.SECTION_HEADING_FONT_SIZE


def test_set_section_heading_font_custom(clear_worksheet, config, excel_helper):
    worksheet = clear_worksheet
    worksheet.append(["Test"])
    cell = worksheet.cell(1, 1)
    excel_helper.set_section_heading_font(
        cell,
        sh_font_name="Arial",
        sh_bold=True,
        sh_font_size=14,
        sh_font_color="FF0000",
    )
    assert cell.font.name == "Arial"
    assert cell.font.size == 14


def test_set_section_heading_font_with_kwargs(clear_worksheet, excel_helper):
    worksheet = clear_worksheet
    worksheet.append(["Test"])
    cell = worksheet.cell(1, 1)
    excel_helper.set_section_heading_font(cell, chart="ignored_kwarg")
    assert cell.font.name == excel_helper.config.SECTION_HEADING_FONT_NAME


def test_set_df_title_font_defaults(clear_worksheet, config, excel_helper):
    worksheet = clear_worksheet
    worksheet.append(["Test"])
    cell = worksheet.cell(1, 1)
    excel_helper.set_df_title_font(cell)
    assert cell.font.name == config.DF_TITLE_FONT_NAME
    assert cell.font.bold == config.DF_TITLE_BOLD
    assert cell.font.size == config.DF_TITLE_FONT_SIZE


def test_set_df_title_font_custom(clear_worksheet, config, excel_helper):
    worksheet = clear_worksheet
    worksheet.append(["Test"])
    cell = worksheet.cell(1, 1)
    excel_helper.set_df_title_font(
        cell,
        df_font_name="Arial",
        df_bold=True,
        df_font_size=14,
        df_font_color="FF0000",
    )
    assert cell.font.name == "Arial"
    assert cell.font.size == 14


def test_set_df_title_font_with_kwargs(clear_worksheet, excel_helper):
    worksheet = clear_worksheet
    worksheet.append(["Test"])
    cell = worksheet.cell(1, 1)
    excel_helper.set_df_title_font(cell, chart="ignored_kwarg")
    assert cell.font.name == excel_helper.config.DF_TITLE_FONT_NAME


def test_set_chart_title_font_none_title(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.title = None
    excel_helper.set_chart_title_font(chart_obj)


def test_set_chart_title_font_default(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.title = "Test Title"
    excel_helper.set_chart_title_font(chart_obj)
    rpr = chart_obj.title.tx.rich.p[0].r[0].rPr
    assert rpr is not None
    assert rpr.sz == config.CHART_TITLE_FONT_SIZE


def test_set_chart_title_font_custom(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.title = "Custom Title"
    excel_helper.set_chart_title_font(
        chart_obj,
        title_font_name="Arial",
        title_font_size=2000,
        title_font_color="FF0000",
        title_font_bold=True,
    )
    rpr = chart_obj.title.tx.rich.p[0].r[0].rPr
    assert rpr.sz == 2000
    assert rpr.b is True


def test_set_chart_title_font_valid_prst_color(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.title = "Prst Color"
    excel_helper.set_chart_title_font(chart_obj, title_font_color="green")
    rpr = chart_obj.title.tx.rich.p[0].r[0].rPr
    assert rpr is not None


def test_set_chart_title_font_empty_title_rich_text(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.title = "Title"
    chart_obj.title.tx.rich.p = []
    with pytest.warns(UserWarning, match="Unable to set chart title"):
        excel_helper.set_chart_title_font(chart_obj)


def test_set_chart_axis_label_font_x_axis(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.x_axis.title = "X Label"
    excel_helper.set_chart_axis_label_font(chart_obj, axis="x")
    rpr = chart_obj.x_axis.title.tx.rich.p[0].r[0].rPr
    assert rpr is not None
    assert rpr.sz == config.AXIS_FONT_SIZE


def test_set_chart_axis_label_font_y_axis(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.y_axis.title = "Y Label"
    excel_helper.set_chart_axis_label_font(chart_obj, axis="y")
    rpr = chart_obj.y_axis.title.tx.rich.p[0].r[0].rPr
    assert rpr is not None


def test_set_chart_axis_label_font_custom(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.x_axis.title = "X Custom"
    excel_helper.set_chart_axis_label_font(
        chart_obj,
        axis="x",
        axis_font_name="Arial",
        axis_font_size=2000,
        axis_font_color="FF0000",
        axis_font_bold=True,
    )
    rpr = chart_obj.x_axis.title.tx.rich.p[0].r[0].rPr
    assert rpr.sz == 2000
    assert rpr.b is True


def test_set_chart_axis_label_font_valid_prst_color(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.x_axis.title = "Prst"
    excel_helper.set_chart_axis_label_font(chart_obj, axis="x", axis_font_color="green")
    rpr = chart_obj.x_axis.title.tx.rich.p[0].r[0].rPr
    assert rpr is not None


def test_set_chart_axis_label_font_no_title_no_crash(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.x_axis.title = None
    chart_obj.y_axis.title = None
    excel_helper.set_chart_axis_label_font(chart_obj, axis="x")
    excel_helper.set_chart_axis_label_font(chart_obj, axis="y")


def test_set_chart_axis_label_font_unknown_axis(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.x_axis.title = "X"
    excel_helper.set_chart_axis_label_font(chart_obj, axis="z")


def test_rotate_xticks(config, excel_helper):
    chart_obj = XlLineChart()
    chart_obj.x_axis.title = "X"
    excel_helper.rotate_xticks(chart_obj, 45)
    assert chart_obj.x_axis.txPr is not None


def test_rotate_xticks_zero(config, excel_helper):
    chart_obj = XlLineChart()
    excel_helper.rotate_xticks(chart_obj, 0)
    assert chart_obj.x_axis.txPr is not None


def test_set_axis_limit_valid(config, excel_helper):
    chart_obj = XlLineChart()
    excel_helper.set_axis_limit(chart_obj.x_axis, (0, 100))
    assert chart_obj.x_axis.scaling.min == 0
    assert chart_obj.x_axis.scaling.max == 100


def test_set_axis_limit_equal_values_warns(config, excel_helper):
    chart_obj = XlLineChart()
    with pytest.warns(UserWarning, match="Invalid axis limits"):
        excel_helper.set_axis_limit(chart_obj.x_axis, (10, 10))
    assert chart_obj.x_axis.scaling.min == 10
    assert chart_obj.x_axis.scaling.max == 10


def test_set_axis_limit_min_greater_than_max_warns(config, excel_helper):
    chart_obj = XlLineChart()
    with pytest.warns(UserWarning, match="Invalid axis limits"):
        excel_helper.set_axis_limit(chart_obj.y_axis, (200, 10))
    assert chart_obj.y_axis.scaling.min == 200
    assert chart_obj.y_axis.scaling.max == 10


def test_set_axis_limit_floats(config, excel_helper):
    chart_obj = XlLineChart()
    excel_helper.set_axis_limit(chart_obj.x_axis, (0.5, 9.5))
    assert chart_obj.x_axis.scaling.min == 0.5
    assert chart_obj.x_axis.scaling.max == 9.5
