"""
Configuration module.

Defines the Config class containing constants for formatting, chart settings,
color schemes, and layout options.
"""


class Config:

    # Setting for section heading
    SECTION_HEADING_FONT_NAME = "Calibri"
    SECTION_HEADING_FONT_SIZE = 11
    SECTION_HEADING_BOLD = True
    # Color in hex code: https://openpyxl.readthedocs.io/en/stable/styles.html
    SECTION_HEADING_FONT_COLOR = "1A1A1A"

    # Dataframe headings
    DF_TITLE_FONT_NAME = "Calibri"
    DF_TITLE_FONT_SIZE = 10
    DF_TITLE_FONT_COLOR = "1A1A1A"
    DF_TITLE_BOLD = True

    # Chart position
    # 'left': NotImplemented
    # 'right': place the chart on the right relative to the data columns
    # 'bottom': place the chart below data
    # 'top': NotImplemented
    CHART_POSITION = "right"

    # Separator: number of empty rows between two sections
    SEPARATOR = 2

    # Separate data from chart
    DATA_CHART_SEPARATOR = 1
    DATA_DATA_SEPARATOR = 1

    # Chart labels settings
    CHART_TITLE_FONT_NAME = "Calibri"
    CHART_TITLE_FONT_SIZE = 1100  # ~11 pt

    # Color using prsClr or srgbClr: https://openpyxl.readthedocs.io/en/stable/api/openpyxl.drawing.colors.html
    CHART_TITLE_FONT_COLOR = "1A1A1A"
    CHART_TITLE_FONT_BOLD = True

    # Axes settings
    MAJOR_UNIT = None  # Default
    # Axis label
    AXIS_FONT_NAME = "Calibri"
    AXIS_FONT_SIZE = 1000  # ~10 pt
    XAXIS_NUMFMT = None
    YAXIS_NUMFMT = None
    # Color using prstClr
    # NOTE check: https://openpyxl.readthedocs.io/en/stable/api/openpyxl.drawing.colors.html
    AXIS_FONT_COLOR = "black"
    AXIS_FONT_BOLD = False

    # Set chart dimensions to have uniform chart dimensions across the whole report
    CHART_WIDTH = None
    CHART_HEIGHT = None

    OPENPYXL_COLORS = False

    # openpyxl default are not appealing
    PRIMARY_COLORS = [
        "0078D7",
        "FD625E",
        # "2FBE8F",
        "73B761",
        "FF8C00",
        "A66999",
        "00BFFF",
        # "1F4E78",
        # "6bFFA1",
        "D83B01",
        "2FBE8F",
        # "2CA02C",
        "FE9666",
        "D8BFD8",
    ]

    # Used to calculate the space occupied by a chart: space = CHART_HEIGHT_SCALE * chart_height
    # NOTE The actual chartsize will depend on operating system and device.
    # Adjust self.config.CHART_HEIGHT_SCALE as you need.
    CHART_HEIGHT_SCALE = 1.9

    # Rotate by 45 degrees
    # XTICKST_ROTATION = -2700000
    # Rotation in degree
    XTICKST_ROTATION = -45

    # Radar chart
    RADAR_STYLE = 26  # Chart style
    RADAR_SHAPE = 4
    RADAR_UNIT_STEPS = None
    RADAR_REF_COLOR = "2A2A2A"
    RADAR_REF_WIDTH = 2
    RADAR_REF_STYLE = "sysDash"

    # Line chart setting
    LINE_WIDTH = 1.5
    # line styles: 'sysDashDotDot', 'solid', 'sysDash', 'lgDash', 'lgDashDotDot', 'dot', 'dashDot', 'dash', 'sysDot', 'sysDashDot', 'lgDashDot'
    LINE_STYLE = "solid"
    LINE_COLOR = "green"
    # Reference line (a line representating a ref value (average, etc) which is drawn differently)
    USE_REF_LINE = True
    LINE_REF_COLOR = "1A1A1A"
    LINE_REF_WIDTH = 1.5
    LINE_REF_STYLE = "sysDash"
    LINE_SMOOTH = True
    MARKER_SIZE = 6
    MARKER_SYMOBLS = [
        "circle",
        "triangle",
        "star",
        "diamond",
        "plus",
        "x",
        "square",
        "dot",
        "dash",
    ]

    # Pie chart
    SHOW_PERCENTAGE = True
    SHOW_CATEGORIES = True
    SHOW_LEGEND_KEY = False
    SHOW_VALUES = False
    SHOW_SERIES_NAME = False

    BAR_CHART_VARYING_COLOR = False

    # Barline chart
    BARLINE_BARCHART_STYLE = 10
    BARLINE_LINECHART_STYLE = 13

    # Legend
    SHOW_LEGEND = True
    # Legend position
    # Possible values: 'r', 't', 'l', 'b', 'tr'
    LEGEND_POSITION = "b"  # default bottom

    # Image inserted into sheet
    IMAGE_WIDTH = None  # pixels
    IMAGE_HEIGHT = None
    # Make up for the rows consumed by the image
    IMAGE_HEIGHT_UNIT = 0.053

    VALID_COLORS = [
        "ltGoldenrodYellow",
        "darkSlateBlue",
        "fuchsia",
        "ltGrey",
        "medTurquoise",
        "moccasin",
        "lavenderBlush",
        "mediumSlateBlue",
        "gold",
        "paleVioletRed",
        "darkGoldenrod",
        # "darkSlateGray",
        "lightGrey",
        "mediumBlue",
        "chocolate",
        "yellow",
        "lightSkyBlue",
        "darkGrey",
        "lightSlateGrey",
        "tomato",
        "turquoise",
        "darkSalmon",
        "mediumTurquoise",
        "mediumSpringGreen",
        "paleGreen",
        "lightGreen",
        "darkGreen",
        "bisque",
        "green",
        "maroon",
        "hotPink",
        "ltSeaGreen",
        "darkBlue",
        "mediumVioletRed",
        "dkGreen",
        "plum",
        # "darkGray",
        "medSpringGreen",
        "wheat",
        "red",
        "aquamarine",
        "dkGrey",
        # "whiteSmoke",
        "yellowGreen",
        # "ltSlateGray",
        "medSlateBlue",
        "dkSlateGrey",
        "ltBlue",
        "lightCoral",
        "steelBlue",
        "paleTurquoise",
        "pink",
        "greenYellow",
        "darkSlateGrey",
        "mistyRose",
        "crimson",
        "darkKhaki",
        "ltGreen",
        "ltSlateGrey",
        "blanchedAlmond",
        "magenta",
        "orange",
        "lightGoldenrodYellow",
        "dkKhaki",
        "cornflowerBlue",
        "coral",
        "limeGreen",
        "mediumPurple",
        "firebrick",
        "ltCoral",
        "ltSkyBlue",
        "mintCream",
        "chartreuse",
        "medBlue",
        "ltPink",
        "snow",
        # "dkSlateGray",
        "ltCyan",
        "deepSkyBlue",
        "thistle",
        "darkSeaGreen",
        "tan",
        "lightPink",
        "paleGoldenrod",
        "rosyBrown",
        "lightSalmon",
        # "lightGray",
        "dkCyan",
        "dkOrange",
        "mediumSeaGreen",
        "lightCyan",
        "orangeRed",
        "mediumAquamarine",
        "aqua",
        "darkMagenta",
        # "ghostWhite",
        "cyan",
        "dkSalmon",
        "lightSeaGreen",
        "brown",
        "skyBlue",
        "ltYellow",
        "dkSeaGreen",
        "darkRed",
        "midnightBlue",
        "orchid",
        "oldLace",
        "violet",
        "dkOrchid",
        "dkBlue",
        "darkCyan",
        "salmon",
        # "navajoWhite",
        "medOrchid",
        "oliveDrab",
        "dimGrey",
        "darkOliveGreen",
        "mediumOrchid",
        "honeydew",
        "linen",
        "peachPuff",
        "springGreen",
        "medVioletRed",
        "cornsilk",
        "papayaWhip",
        "forestGreen",
        "dkSlateBlue",
        "lightYellow",
        "purple",
        "teal",
        # "white",
        "dkRed",
        "medSeaGreen",
        "seaGreen",
        "seaShell",
        # "gray",
        "dkViolet",
        "navy",
        "lawnGreen",
        "slateBlue",
        # "dimGray",
        "dodgerBlue",
        "lightBlue",
        "dkMagenta",
        "royalBlue",
        # "floralWhite",
        "powderBlue",
        "grey",
        "aliceBlue",
        "silver",
        "gainsboro",
        "dkOliveGreen",
        "indigo",
        "lightSteelBlue",
        "khaki",
        "ltSteelBlue",
        "cadetBlue",
        "black",
        "ltSalmon",
        "darkOrchid",
        "saddleBrown",
        "goldenrod",
        "darkViolet",
        # "lightSlateGray",
        "darkTurquoise",
        "peru",
        "beige",
        # "ivory",
        "medPurple",
        "slateGrey",
        "olive",
        # "antiqueWhite",
        "blueViolet",
        "blue",
        "medAquamarine",
        # "ltGray",
        "azure",
        "dkGoldenrod",
        "darkOrange",
        "indianRed",
        # "dkGray",
        "sienna",
        # "slateGray",
        "lime",
        "dkTurquoise",
        "sandyBrown",
        "lemonChiffon",
        "deepPink",
        "burlyWood",
        "lavender",
    ]
