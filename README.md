<div>
<p align="center">
<img src="data/SQL2Excel_logo.png" style="display: block; margin: auto; width: 250px; height:250px;">
</p>
<p align="center">A Python library for exporting SQL queries and/or Pandas DataFrames to Excel, with support for chart generation and data visualization.</p>
</div>

# SQL2Excel

`SQL2Excel` exports result sets from SQL queries and/or Pandas DataFrames into Excel including chart generation and data visualization with a _strong focus on automating the generation of reports from SQL and Python scripts_. `SQL2Excel`
simplifies the process of exporting data from SQL/Python to Excel. It supports various chart types and allows for basic chart customization. Under the hood, `SQL2Excel` uses [Openpyxl](https://openpyxl.readthedocs.io/en/stable/).

# Table of Contents

- [Why Use SQL2Excel](#why-use-sql2excel)
- [User Guide](#user-guide)
  - [Installation](#installation)
  - [Generating Reports from SQL Script](#generating-reports-from-sql-script)
    <!-- - [Writing Data](#writing-data)
    - [Writing Data and Generating Charts](#writing-data-and-generating-charts)
    - [Chart Customization](#chart-customization)
    - [Pivoting Result Set](#pivoting-result-set)
    - [Generating Charts from Specific Columns](#generating-charts-from-specific-columns)
      - [Using `data_column_start` and `data_column_end`](#using-data_column_start-and-data_column_end)
      - [Using `data_columns`](#using-data_columns) -->
  - [Generating Reports from Queries Executed Within Python](#generating-reports-from-queries-executed-within-python)
    <!-- - [QueryConfig](#queryconfig)
    - [Passing Named Parameters to SQL Queries](#passing-named-parameters-to-sql-queries)
    - [Passing Positional Parameters to SQL Queries](#passing-positional-parameters-to-sql-queries)
    - [Data and Chart Options](#data-and-chart-options)
    - [Running QueryConfig](#running-queryconfig) -->
  - [Generating Reports from DataFrames](#generating-reports-from-dataframes)
    <!-- - [Writing Data](#writing-data)
    - [Writing Data and Generating Charts](#writing-data-and-generating-charts) -->
    <!-- - [Writing Data Side-by-Side](#writing-multiple-dataframes-side-by-side) -->
  - [Matplotlib Figures](#matplotlib-figures)
  - [External Images](#external-images)
  - [Charts](#charts)
    <!-- - [Area Chart](#area-chart)
    - [Bar Chart](#bar-chart)
    - [Barline Chart](#barline-chart)
    - [Bubble Chart](#bubble-chart)
    - [Image Chart](#image-chart)
    - [Line Chart](#line-chart)
    - [Pie Chart](#pie-chart)
    - [Radar Chart](#radar-chart)
    - [Scatter Chart](#scatter-chart)
    - [Stackedbar Chart](#stackedbar-chart) -->
- [Examples](#examples)
- [Miscellaneous](#miscellaneous)
  - [Controlling Result Placement](#controlling-result-placement)
  - [Overriding Defaults](#overriding-defaults)
  - [Featured Options](#featured-options)
  - [Colors](#colors)
- [Running Test](#running-tests)
- [Contribution](#contribution)
- [TO-DO](#to-do)

# Why Use SQL2Excel

The main reasons why you should consider using `SQL2Excel` despite the existence of `Openpyxl` are:

- Integration with SQL
  - `SQL2Excel` executes your SQL script and exports all results to Excel including chart generation.
  - `SQL2Excel` offers support for exporting data from parameterized queries. Hence, increases the efficiency in carrying out repetitive tasks.
- Support for Pandas DataFrames

  - `SQL2Excel` exports DataFrames to Excel. While `Pandas` supports exporting dataframe to Excel, `SQL2Excel` offers more than just writing the dataframe to Excel (e.g. can generate Excel charts from dataframes).

- `Matplotlib` and externally generated images
  - Inserting `Matplotlib` figures and externally generated images is straightforward in `SQL2Excel`. This is useful if you require advanced charts that cannot be generated in Excel.
- Simple API
  - `SQL2Excel` offers simple API for data export and chart generation.

#### Story of `SQL2Excel`

`SQL2Excel` stems from the a project at [CREST](https://www0.sun.ac.za/crest/) to profile the scientific performance of the African countries across 52 scientific disciplines using the [Web of Science](https://clarivate.com/products/scientific-and-academic-research/research-discovery-and-workflow-solutions/webofscience-platform/) database (resulting in generating 1560 reports, each with several performance indicators and different visulaizations).

#### What SQL2Excel is Not

`SQL2Excel` is not meant as library for reading/writing into Excel and it is not suitable for fine customizations of Excel. **Its purpose is to improve productivity, automate repetitive tasks, and simplify data export and chart creation.**
If you want more control over excel from within Python, use `openpyxl` instead.

# User Guide

## Installation

Currently, `SQL2Excel` is not available on PyPI. To use `SQL2Excel`, you will need to clone this repository and install it _locally_ using `pip`.

- Clone the repository

```bash
git clone <repository-url>
```

- Navigate to the root directory of the project (the directory which contains `README.md`) and create a virtual environment (`.env`)

```python
python -m venv .env
```

- Activate the environment

```bash
source .env/bin/activate
```

<!-- - Install the requirements

```bash
pip install -r requirements.txt
``` -->

- Install `SQL2Excel` and its dependencies locally (note the dot `.` after `pip install`)

```bash
pip install .
```

- Verify the Installation

```bash
python -c "import sql2excel; print(f'version: {sql2excel.__version__}')";
python -c "from sql2excel.chart import Chart; print(Chart.__name__)"
```

- Optional, run the tests; see [running tests](#running-tests).

**You can now import `SQL2Excel` in your project as you would normally do.** For example:

```python
from sql2excel.chart import BarChart
```

## Generating Reports from SQL Script

To generate a report from SQL script, you need to do two steps:

- Annotate the SQL script by including _directives_ (special comments).
- Create a `Report` object and call its `generate` method.

These two steps will be explained below.

### Writing Data

To write the result of a SQL query into an excel, add `-- chart` directive as a comment anywhere above the query.

```sql
-- chart
SELECT
    name,
    MIN(rental_rate),
    AVG(rental_rate),
    MAX(rental_rate)
FROM
    category c
    INNER JOIN film_category fc ON c.category_id = fc.category_id
    INNER JOIN film f ON fc.film_id = f.film_id
GROUP BY
    1
ORDER BY
    3 DESC;
```

The result in the Excel workbook is as shown below:

![writing_data](data/writing_data.png)

### Writing Data and Generating Charts

To write query result set and generate a chart from the data, specify the _chart type_ by placing a directive anywhere above the SQL code as follows `-- chart=chart_type` or `-- chart:chart_type` (using `=` or `:` as an option separator) where `chart_type`
is any of the supported charts listed below. For example, this comment
`-- chart=line` will generate a line chart from the query result. Currently, the
supported charts are shown below:

<table>
  <tr>
    <th>Chart</th>
    <th>Directive</th>
  </tr>
    <tr>
    <td>area</td>
    <td>-- chart=area</td>
  </tr>
    <tr>
    <td>bar</td>
    <td>-- chart=bar</td>
  </tr>
  <tr>
    <td>barline</td>
    <td>-- chart=barline</td>
  </tr>
    <tr>
    <td>bubble</td>
    <td>-- chart=bubble</td>
  </tr>
  <tr>
    <td>chart</td>
    <td>-- chart=chart</td>
  </tr>
  <tr>
    <td>line</td>
    <td>-- chart=line</td>
  </tr>

  <tr>
    <td>pie</td>
    <td>-- chart=pie</td>
  </tr>
  <tr>
    <td>radar</td>
    <td>-- chart=radar</td>
  </tr>
  <tr>
    <td>scatter</td>
    <td>-- chart=scatter</td>
  </tr>
  <tr>
    <td>stackedbar</td>
    <td>-- chart=stackedbar</td>
  </tr>
</table>

For example, the result set of the query below will be exported to Excel and a bar chart will be generated from the result set.

```sql
-- chart=bar
SELECT
    name AS category,
    count(DISTINCT rental_id) AS rental_count
FROM
    category c
    INNER JOIN film_category fc ON c.category_id = fc.category_id
    INNER JOIN film f ON fc.film_id = f.film_id
    INNER JOIN inventory i ON f.film_id = i.film_id
    INNER JOIN rental r ON i.inventory_id = r.rental_id
GROUP BY
    1
ORDER BY
    2 DESC;
```

![basic chart](data/basic_chart.png)

### Chart Customization

`SQL2Excel` offers options for basic customization. Customization options can be provided as directives above the SQL query in one line or spread over multiple lines. An example is shown below

```sql
-- chart=line, width=25, height=12
-- section_heading=Drawing a chart with some customization, title=Rental count
-- xlabel=Date, ylabel=Rental Count
-- line_width=1.5, line_style=sysDash, marker_symbol=circle, marker_size=8, line_color=A66999
SELECT
    rental_date :: date AS "Rental Date",
    COUNT(DISTINCT rental_id) AS "Rental Count"
FROM
    rental
GROUP BY
    1
ORDER BY
    1;
```

![customize_chart](data/customizing_chart.png)

### Pivoting Result Set

Pivoting data to convert it from long format to wide format typically requires writing cumbersome SQL code or using external extensions in some DBMS. `SQL2Excel` simplifies pivotting result set. You need to decorate your query with three parameters:

- index: The column to use as the index in the pivoted table. Each value in the index column will have one row in the pivoted table.
- columns: The values in this column will be used as the columns in the pivoted table.
- values: The column to use as the values in the pivoted table.

```sql
-- chart: bar, section_heading=Query result with pivoting data
-- index=month_year, columns=category_name, values=total_payment
-- show_legend=true, ylabel=Monthly spending, rotation=0, chart_position=bottom
SELECT
    extract(
        'month'
        FROM
            p.payment_date
    ) || '-' || extract(
        'year'
        FROM
            p.payment_date
    ) AS month_year,
    c.name AS category_name,
    SUM(p.amount) AS total_payment
FROM
    payment p
    JOIN rental r ON p.rental_id = r.rental_id
    JOIN inventory i ON r.inventory_id = i.inventory_id
    JOIN film f ON i.film_id = f.film_id
    JOIN film_category fc ON f.film_id = fc.film_id
    JOIN category c ON fc.category_id = c.category_id
WHERE
    c.name IN ('Family', 'Action', 'Sports', 'Music')
GROUP BY
    1,
    2
ORDER BY
    1,
    2;
```

<div>
<p>The result <strong>withtout pivoting</strong> is shown below</p>
<img src="data/without_pivot.png">
</div>

<div>
<p>The result <strong>with pivoting</strong> is shown below</p>
<img src="data/with_pivot.png">
</div>

### Generating Charts from Specific Columns

By default, _all columns returned by a query are written into Excel_, and all columns are used to generate the chart.
_The first column must be the category (x-axis)_. If your query does not meet this requirement, rearrange the columns as such. It is possible to write all columns into an Excel sheet but use _specific columns_ to generate the chart. There are two ways to do this:

- Use `data_column_start` and `data_column_end` if the columns to be used in the chart are _consecutive_
- Use `data_columns` if the columns to be used in the chart are not consecutive.

#### Using `data_column_start` and `data_column_end`

`data_column_start` and `data_column_end` allow you to select specific _continguous_ columns in the result set for the chart.

- `data_column_start`: the starting column index for the chart data.
- `data_column_end`: the ending column index for the chart data.

**Note:** indexes start from 1 (first column or first row depending on the context).

For example, in the query below, all three columns will be written. However, the bar chart will only
be genereated from the first column (x-axis) and the third column (amount) because of `data_column_start:3`.
In this case, you could have dropped `data_column_end:3` because we only drawing one column.

```sql
--chart:bar, data_column_start:3, data_column_end:3
-- section_heading: Selecting consecutive columns using data_column_start and data_column_end
--title: Customer Spending, ylabel: Rental amount, vary_color: True
SELECT
    c.first_name || ' ' || c.last_name AS full_name,
    c.customer_id,
    SUM(p.amount) AS total_amount
FROM
    customer c
    JOIN payment p ON c.customer_id = p.customer_id
WHERE
    c.customer_id IN (1, 2, 3, 4, 5)
GROUP BY
    1,
    2
ORDER BY
    3 DESC;
```

![data_column_start](data/column_start_end.png)

#### Using `data_columns`

You can be more explicit by listing specific columns to include in the chart using `data_columns`. This is useful when the columns to be used in the chart are _not consecutive_. `data_columns` is a comma-separated list of column indices to use for the chart.

```sql
-- chart:bar
-- Use the 2nd and 4th columns in the bar chart
-- data_columns=[2, 4], section_heading=Selecting non-consecutive columns
SELECT
    name,
    min(rental_rate),
    avg(rental_rate),
    max(rental_rate)
FROM
    category c
    INNER JOIN film_category fc ON c.category_id = fc.category_id
    INNER JOIN film f ON fc.film_id = f.film_id
GROUP BY
    1
ORDER BY
    3 DESC;
```

![data_column_start](data/data_columns.png)

### Organizing Query Results into Separate Sheets

By default, all query results are exported into the same sheet within the Excel workbook. However, in many cases, you may need to separate results into different sheets based on logical groupings.

`SQL2Excel` allows you to organize related queries into separate sheets using the `--sheetname=[sheet name]` directive. Once this directive is encountered, all subsequent query results are written to the specified sheet until another `--sheetname` directive is encountered.

If the sheet name already exists, `SQL2Excel` appends results to that sheet instead of creating a new one. For example, the queries below will be written into two sheets: "sales" and "cost". The first and third query will be written into "sales" worksheet and the second query will be written to the "cost" worksheet

```sql
-- sheetname=sales
--chart
select *
from sales_table
where product='Laptop';

-- sheetname=cost
--chart
select *
from cost_table;

-- sheetname=sales
--chart
select *
from sales_table
where product='Phone';
```

**The recommneded convention is to place `--sheetname` directive in its own line.**

### Generating the Excel Report

Once you have annotated your SQL script with the appropriate directives (SQL comments as explained above), you need to create a Report object and call its generate method to execute the queries and write the results into an Excel workbook.

```python
from sql2excel.parser import parse_sql_file
from sql2excel.report import Report
# Database parameters stored in a local file in your machine
from db_params import connection_string

# An example of connection string for Postgres:
# 'postgresql://username:password@host:port/database_name'

# NOTE connection_string is the database URL of SQLAlchemy


# Provide the path to the SQL script
query_configs = parse_sql_file(
    "path/to/sql/my_script.sql"
)

# Create a report object
report = Report(connection_string=connection_string)

# Generate the report and save the results to a file
report.generate(
    query_configs, fname="path/to/save/excel/my_report.xlsx"
)
```

**NOTE: Install the database driver for your DBMS.** For PostgreSQL, `psycopg2` is already installed.

[Back to Top](#table-of-contents)

## Generating Reports from Queries Executed Within Python

If you work with SQL from within Python, you can still export the results and generate charts using `SQL2Excel`. For this purpose you need to create `QueryConfig` object (see below), then create a `Report` object and invoke its `generate` method (see below).

### QueryConfig

`QueryConfig` defines everything `SQL2Excel` needs to run your query and export data to Excel. To construct a `QueryConfig` object you need to provide the following arguments:

- SQL Query (`sql`): This is the SQL query you want to run.
- SQL Parameters (`sql_params`): If your query needs some parameters to run (e.g., a specific date or range of numbers, etc), you can pass these as named parameters or positional parameters.
- Excel Options: _keyword arguments_ to control data export and visualization such as chart type, labels or titles for the charts

#### Basic QueryConfig

An example of a basic QueryConfig object is shown below

```python
from sql2excel.sqlexec import QueryConfig

# Query with no parameters
query = """SELECT rental_date::date, COUNT(DISTINCT rental_id) AS rental_count
           FROM rental
           GROUP BY 1
           ORDER BY 1;"""

# Create a QueryConfig object
query = QueryConfig(sql=query, chart="line")
```

### Passing Named Parameters to SQL Queries

`SQL2Excel` uses `SQLAlchemy` to connect and query databases. Therefore, you can pass parameters to
your query the same way you would if you used `SQLAlchemy`. In the query below `rental_date_param` is a
parameter that will be passed to the query (note the `:` before the parameter).

```python
# Query with name
query = """SELECT rental_date::date, COUNT(DISTINCT rental_id) AS rental_count
           FROM rental
           WHERE rental_date > :rental_date_param
           GROUP BY 1
           ORDER BY 1;"""

query = QueryConfig(
    sql=query,
    # You pass named parameters to query by specifying `sql_params` as a dict
    sql_params={"rental_date_param": "2005-07-31"},
    chart="bar",
)
```

### Passing Positional Parameters to SQL Queries

You can pass positional parameters to your SQL queries by marking them with `?` and providing the values of the parameters as a `Sequence` (e.g. list, tuple, etc). Note that positional parameters are processed in the same order they are provided.

```python
# positional parameter
query = """SELECT length AS film_length, COUNT(DISTINCT film_id) AS film_count
           FROM film
           WHERE length BETWEEN ? AND ?
           GROUP BY length
           ORDER BY film_length;"""

# Note that sql_params is a tuple this time.
query = QueryConfig(sql=query, sql_params=(100, 110), chart="chart")
```

**Note:** If you want to pass a one-element tuple, remember to include the comma at the end as `(my_param_value,)`.

**Note:** SQLAlchemy 2.0 does not support positional parameters [(here)](https://github.com/sqlalchemy/sqlalchemy/issues/5178)
Positional parameters are easier to work with even though named
parameters are more verbose and easier to understand. The issue
above suggests using `exec_driver_sql`. However, passing parameters
this way is not DB-agnostic and will require manual adjustment.
`SQL2Excel` uses an alternative approach, which is database-agnostic, to handle
positional parameters.

### Data and Chart Options

As mentioned, any options related to Excel should be passed as _keyword arguments_.
To write data without generating chart, use the option `chart="chart"`. An example
of chart customization is shown below

```python
query = """SELECT length AS film_length, COUNT(DISTINCT film_id) AS film_count
           FROM film
           WHERE length BETWEEN ? AND ?
           GROUP BY length
           ORDER BY film_length;"""

query = QueryConfig(sql=query,
                    sql_params=(90, 120),
                    chart="scatter",
                    xlabel="Film Length",
                    ylabel="Number of Films",
                    title="Film Length Distribution"
                  )
```

### Generating Report from QueryConfig

Once you define your `QueryConfig` object, you can export the result set of `QueryConfig` by creating a `Report` object and call its `generate` method as in the example below.

```python
from db_params import connection_string

# An example of connection string:
# 'postgresql://username:password@localhost:5432/database_name'

query_config = .... # Create a QueryConfig object as explained above

report = Report(connection_string=connection_string)

report.generate(
    query_config, fname="path/to/excel/file.xlsx"
)
```

For long reports involving multiple queries, you can also send a list of `QueryConfig` objects to `generate` method of `Report` class.

```python
query_config_list = []

query_config1 = .... # Create a QueryConfig object

query_config_list.append(query_config1)

query_config2 = .... # Create another QueryConfig object

query_config_list.append(query_config2)

# Generate a report from both query_config1 and query_config2
report.generate(
    query_config_list, fname="path/to/excel/file.xlsx"
)
```

[Back to Top](#table-of-contents)

## Generating Reports from DataFrames

`SQL2Excel` supports exporting data and generating Excel visualization from DataFrames.
In fact, `SQL2Excel` works with DataFrame under the hood.

**Note:** Pivoting data (converting data from long format data to wide format data) is only supported
if you use SQL scripts or QueyConfig as explained above. If you are working with DataFrames direclty,
pivot your data beforing calling `write_dataframe` or `plot` methods of `Chart` class.

### Writing Data

To export DataFrame to Excel **without** generating charts,
you need to create an instance of the `Chart` class and call its `write_dataframe` method.
You also need to create a `Workbook` using `Openpyxl` package and write the data into a sheet. See the example below.

```python
# Use openpyxl to create Workbook (Excel file) and save it
import openpyxl as xl
from sql2excel.chart import Chart

df = ... # Create DataFrame

# Create a workbook
wb = xl.Workbook()

# Choose the active sheet
ws = wb.active

# Create a chart object
chart = Chart()

# Write dataframe
chart.write_dataframe(df, ws)

# Save the result
wb.save('path/to/file.xlsx')
```

### Writing Data and Generating Charts

You can write data and generate a chart by creating an instance of the chart class of your choice and call its `plot` method. You can either place the chart to the right of the data or below the data by setting `chart_position` to `'right'` or `'bottom'` respectively. Default is `'right'`.

```python
import openpyxl as xl
from sql2excel.chart import LineChart

df = ... # Create DataFrame

# Create a workbook
wb = xl.Workbook()

# Choose the active sheet
ws = wb.active

# Create dummy data
desktop = ... # sequence of numbers (deskop sales)
laptop = ...
tablet = ...
phone = ...

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

# NOTE: The baseline is detected and formatted differently
# A baseline is a reference line with a constant value.
chart = LineChart()
chart.plot(df, ws)

chart = LineChart()
chart.plot(
    df,
    ws,
    section_heading="Monthly sales (selecting specific columns - not consecutive)",
    title="Monthly sales",
    data_columns=[2, 4, 5], # which column to include in the chart
    line_width=3,
    line_style="sysDash",
    marker_symbol="circle",
    marker_size=8,
    # row_start=40, # You can specify the row where you want to include the figure
    column_start=3,  # the column where to start writing the data
    xlabel="Month",
    ylabel="Sales",
    width=25, # chart width
    height=10, # chart height
    chart_position="bottom", # chart position relative to the data
    rotation=30,
    # show_legend=False,  # You can disable legend
    legend_position="t",  # legend position: 'r', 't', 'l', 'b'
)
```

![line chart](data/line_chart.png)

Notice that we start from Column C (because `column_start=3`). By default, the x-axis ticks are rotated by 45 degree. You can set `rotation=0` if no rotation is needed.

`LineChart` detects a column that is constant (e.g. average) and plot it differently (e.g. in the above chart, you can see the column Baseline is plotted differently). You can customize this line from [config.py](sql2excel/config.py). To disable this feature, you can either set `use_ref_line` to `False` when invoking `plot` or override the default configuration and set `USE_REF_LINE` to `False`. See [Overriding Defaults](#overriding-defaults) for more information on overriding the default configurations. Refer to [line_chart.py](examples/line_chart.py) for examples of how to do this.

[Back to Top](#table-of-contents)

## Matplotlib Figures

If your report requires advanced `Matplotlib` figures, you can use the `ImageChart` class to insert figures generated by `Matplotlib` into Excel. A code snippet illustrating this is given below.

```python
from sql2excel.chart import ImageChart

# To insert the figure, use ImageChart
chart = ImageChart()

# Method 1: Using plt
fig = plt.figure()
plt.plot(df.X, df.Y)
plt.xlabel("X")
plt.ylabel("Y")
plt.title("Using plt")
plt.close(fig)

# Default insertion (DataFrame will NOT be written)
chart.add_image(fig, ws, section_heading="Inserting Matplotlib figures (using plt)")

# Method 2: Using axes
fig, ax = plt.subplots()
ax.plot(df.X, df.Y)
plt.xlabel("X")
plt.ylabel("Y")
plt.title("Using Axes")
plt.close(fig)

# Default insertion again
chart.add_image(fig, ws, section_heading="Inserting Matplotlib figures (using axes)")

# When you pass DataFrame, it will be written as well
chart.add_image(
    fig,
    ws,
    df=df,
    section_heading="When you pass DataFrame, it will be written as well",
)
```

![Matplotlib](data/matplotlib.png)

Refer to [matplotlib2excel.py](examples/matplotlib2excel.py) for the code and to [matplotlib.xlsx](data/matplotlib.xlsx) to see the output.

[Back to Top](#table-of-contents)

## External Images

The `ImageChart` class allows you to insert images from your local device by providing the absolute path to the image file.
This is ideal for adding figures generated by external tools into Excel.
The class also enables the inclusion of the data used to create the figure.

```python
from sql2excel.chart import ImageChart


# Read data
df = pd.read_csv(
    "data/Africa_collab_2013_2022.csv", delimiter=";", keep_default_na=False
)

# Insert into excel
chart = ImageChart()
# Adjust width and height while preserving aspect ratio
chart.add_image("data/Africa_collab_2013_2022.png", ws, df=df, width=1425, height=552)
```

![image2excel](data/image2excel.png)

**NOTE:** when inserting images, you need to set `width` and `height` as shown above. Otherwise, your image can appear as a very small rectangle that is hardly visible.

## Charts

`SQL2Excel` supports several charts. Examples of how to use each chart in provided below.

### Area Chart

Refer to [Area Chart examples](examples/area_chart.py) and to [Excel output](data/area_chart.xlsx) for the generated Excel workbook.

### Bar Chart

Refer to [Bar Chart examples](examples/bar_chart.py) and to [Excel output](data/bar_chart.xlsx) for the generated Excel workbook.

### Barline Chart

A barline chart is not a standard chart in `Openpyxl` but is custom-built within `SQL2Excel` by combining a bar chart with a line chart, each with its own axis. This chart is ideal for presenting related data in one view, making it easier to analyze both absolute values and proportions. For instance, you can use a barline chart to show the number of customers in the top 10 countries with bars, while the percentage share of these countries out of the total customers is represented by a line. To produce a barline chart, decorate your sql query with `chart=barline`.

```sql
-- chart=barline
SELECT
    ctr.country,
    count(DISTINCT c.customer_id) AS customer_count,
    round(
        100.0 * count(DISTINCT c.customer_id) / (
            SELECT
                count(DISTINCT customer_id)
            FROM
                customer
        ),
        2
    ) AS percentage_share_customer
FROM
    country ctr
    LEFT JOIN city ON ctr.country_id = city.country_id
    LEFT JOIN address a ON city.city_id = a.city_id
    LEFT JOIN customer c ON a.address_id = c.address_id
GROUP BY
    1
ORDER BY
    2 DESC
LIMIT
    10;
```

![Barline](data/barline_chart.png)

Refer to [Barline Chart](examples/barline_chart.py) for the code and to [barline_chart.xlsx](data/barline_chart.xlsx) for the generated Excel workbook.

### Bubble Chart

Refer to [Bubble Chart examples](examples/bubble_chart.py) and to [Excel output](data/bubble_chart.xlsx) for the generated Excel workbook.

### Image Chart (Insert Image into Excel)

Refer to [Image Chart examples](examples/image2excel.py) and to [Excel output](data/image_chart.xlsx) for the generated Excel workbook.

### Line Chart

Refer to [Line Chart examples](examples/line_chart.py) and to [Excel output](data/line_chart.xlsx) for the generated Excel workbook.

### Pie Chart

Refer to [Pie Chart examples](examples/pie_chart.py) and to [Excel output](data/pie_chart.xlsx) for the generated Excel workbook.

### Radar Chart

Refer to [Radar Chart examples](examples/radar_chart.py) and to [Excel output](data/radar_chart.xlsx) for the generated Excel workbook.

### Scatter Chart

Refer to [Scatter Chart examples](examples/scatter_chart.py) and to [Excel output](data/scatter_chart.xlsx) for the generated Excel workbook.

### Stackedbar Chart

A Stackedbar chart is not a standard chart in `Openpyxl` but is custom-built within `SQL2Excel`. It is bar chart with data series stacked on top of each other.

Refer to [Stackedbar Chart examples](examples/stackedbar_chart.py) for code example and to [Excel output](data/stackedbar_chart.xlsx) for the generated Excel workbook.

[Back to Top](#table-of-contents)

# Examples

Refer to [examples](examples) for further code examples.

[Back to Top](#table-of-contents)

# Miscellaneous

## Controlling Result Placement

In `SQL2Excel`, the starting row for writing results is dynamically determined based on the existing data and charts in the current sheet.
You can customize the spacing between two results set or between a result set and chart in the `sql2excel.Config` class:

- `SEPARATOR`: Specifies the number of empty rows to insert between two result sets. It is applied **after** the row in which the new result will be written is calculated automatically depending on the existing data in the current sheet. For example, if the first result has 10 rows and you set the `SEPARATOR=2`, the second result will start at row 12.
- The `CHART_HEIGHT_SCALE`: defines the number of rows the chart occupies, calculated as `rows = CHART_HEIGHT_SCALE * chart_height`.

**Note:** By default, `CHART_HEIGHT_SCALE` is set to `1.9`. However, the actual chart size will depend on operating system and device. If the default value does not meet your needs, adjust as required in your OS and device.

If required, the starting position of the result set can be determined manually with the arguments `row_start` and/or `column_start` of the `plot` or `write_dataframe` method of `Chart` class (and its subclasses).

## Overriding Defaults

For consistency, some settings (e.g. font properties, color theme, etc) are set globally. There are two ways to override these defaults:

- Provide keyword arguments to `plot` or `write_dataframe` methods of the `Chart` class (and its subclasses).
- Create a `Config` object and override its attributes as needed.

The first method does not cover all attributes declared in `Config` class. The second approach is the preferred one. An example is shown below

```python
from sql2excel.config import Config
from sql2excel.chart import BarChart

# Create a config object
config = Config()

# Change font size of the chart title to 11 pt (default is 11 pt)
config.CHART_TITLE_FONT_SIZE = 1400

# Use normal font weight
config.CHART_TITLE_FONT_BOLD = False

# Pass config to chart as a KEYWORD argument
chart = BarChart(config=config)

# You can reuse `config` with other charts to apply the same settings across the report
```

## Featured Options

<table>
  <tr>
    <th>Option</th>
    <th>Description</th>
  </tr>
  <tr>
    <td><b>Data Writing</b></td>
    <td></td>
  </tr>
  <tr>
    <td>write_dataframe</td>
    <td>A method of `Chart` (and its subclasses) to write a result set (SQL/DataFrame) into an Excel sheet.</td>
  </tr>
  <tr>
    <td>write_dataframes_side_by_side</td>
    <td>A method of `Chart` (and its subclasses) to write multiple DataFrames side by side in an Excel sheet.</td>
  </tr>
  <tr>
    <td><b>Chart Generation</b></td>
    <td></td>
  </tr>
  <tr>
  <td>plot</td>
   <td>A method of `Chart` (and its subclasses) to plot charts</td>
  </tr>
 <tr>
    <td><b>Manual Placement of Resultset</b></td>
    <td></td>
  </tr>
  <tr>
    <td>column_start</td>
    <td>Sets the starting column for writing data. Defaults to 1.</td>
  </tr>
  <tr>
    <td>row_start</td>
    <td>Sets the starting row for writing data. The default is determined dynamically depending on the existing data in the sheet.</td>
  </tr>
  <tr>
    <td>chart_position</td>
    <td>Specifies the position of the chart relative to data. Defaults to 'right' (the chart is written next to the data). Possible values: 'right' and 'bottom'.</td>
  </tr>
  <tr>
    <td><b>Data for Charts</b></td>
    <td></td>
  </tr>
  <tr>
    <td>data_columns</td>
    <td>Specifies the columns to be used for chart data. Default all columns will be used and the first column is assumed the category (the x-axis). Useful when the specified columns are not consecutive.</td>
  </tr>
  <tr>
    <td>data_column_start</td>
    <td>Specifies the starting column for data in the chart. Useful when the specified columns are consecutive.</td>
  </tr>
  <tr>
    <td>data_column_end</td>
    <td>Specifies the end column for data in the chart. Useful when the specified columns are consecutive.</td>
  </tr>
  <tr>
    <td><b>Chart Formatting</b></td>
    <td></td>
  </tr>
  <tr>
    <td>chart_shape</td>
    <td>Sets the shape of the chart (same as in Openpyxl).</td>
  </tr>
  <tr>
    <td>chart_style</td>
    <td>Sets the style of the chart (same as in Openpyxl).</td>
  </tr>
  <tr>
    <td>legend_position</td>
    <td>Sets the position of the chart legend.</td>
  </tr>
  <tr>
    <td>show_legend</td>
    <td>Determines whether to show the chart legend.</td>
  </tr>
  <tr>
    <td>smooth</td>
    <td>Sets whether the lines are smoothed or not (relevant for LineChart only).</td>
  </tr>
  <tr>
    <td>height</td>
    <td>Sets the height of the chart.</td>
  </tr>
  <tr>
    <td>width</td>
    <td>Sets the width of the chart.</td>
  </tr>
  <tr>
    <td>border_line_color</td>
    <td>Sets the border line color for the chart series.</td>
  </tr>
  <tr>
    <td>line_color</td>
    <td>Sets the line color for the chart series.</td>
  </tr>
  <tr>
    <td>line_style</td>
    <td>Sets the line style for the chart series.</td>
  </tr>
  <tr>
    <td>line_width</td>
    <td>Sets the line width for the chart series.</td>
  </tr>
  <tr>
    <td>nofill</td>
    <td>Prevents filling the chart series with colors.</td>
  </tr>
  <tr>
    <td>marker_size</td>
    <td>Sets the size of the markers on the chart series.</td>
  </tr>
  <tr>
    <td>marker_symbol</td>
    <td>Sets the symbol of the markers on the chart series.</td>
  </tr>
  <tr>
    <td><b>Chart Title</b></td>
    <td></td>
  </tr>
    <tr>
    <td>title</td>
    <td>Sets the title of the chart.</td>
  </tr>
    <tr>
    <td>title_font_name</td>
    <td>Sets the font name of chart title..</td>
  </tr>
    <tr>
    <td>title_font_size</td>
    <td>Sets the font size of chart title..</td>
  </tr>
  <tr>
    <td>title_font_color</td>
    <td>Sets the color of the font of chart title.</td>
  </tr>
  <tr>
    <td>title_font_bold</td>
    <td>Whether to set the chart title's font to bold.</td>
  </tr>
  <tr>
    <td><b>Axis Settings</b></td>
    <td></td>
  </tr>
  <tr>
    <td>x_log_base</td>
    <td>Sets the logarithmic base for the x-axis.</td>
  </tr>
  <tr>
    <td>y_log_base</td>
    <td>Sets the logarithmic base for the y-axis.</td>
  </tr>
  <tr>
    <td>xlim</td>
    <td>Sets the limits for the x-axis.</td>
  </tr>
  <tr>
    <td>ylim</td>
    <td>Sets the limits for the y-axis.</td>
  </tr>
  <tr>
    <td>xlabel</td>
    <td>Sets the label for the x-axis.</td>
  </tr>
  <tr>
    <td>ylabel</td>
    <td>Sets the label for the y-axis.</td>
  </tr>
  <tr>
    <td>rotation</td>
    <td>Sets the rotation of the x-axis tick labels.</td>
  </tr>
  <tr>
    <td>y_orientation</td>
    <td>Sets the orientation of the y-axis.</td>
  </tr>
  <tr>
    <td>yaxis_major_unit</td>
    <td>Sets the major unit for the y-axis.</td>
  </tr>
  <tr>
    <td>axis_font_name</td>
    <td>Sets the the font name of the xlabel and ylabel.</td>
  </tr>  
  <tr>
    <td>axis_font_size</td>
    <td>Sets the the font size of the xlabel and ylabel.</td>
  </tr>  
  <tr>
    <td>axis_font_color</td>
    <td>Sets the the font color of the xlabel and ylabel.</td>
  </tr>    
  <tr>
    <td>axis_font_bold</td>
    <td>Whether to set the the font of the xlabel and ylabel to bold.</td>
  </tr>    
  <tr>
    <td>xtick_label_position</td>
    <td>Placement of xticks.</td>
  </tr> 
  <tr>
    <td><b>Miscellaneous</b></td>
    <td></td>
  </tr>
  <tr>
    <td>add_image (ImageChart)</td>
    <td>Adds an image to the sheet. Useful for including Matplotlib figures and/or externally generated figures.</td>
  </tr>
  <tr>
    <td>section_heading</td>
    <td>A title to add in the cell above the DataFrame.</td>
  </tr>
</table>

## Colors

### Color Scheme

`SQL2Excel` overrides the default color scheme of `Openpyxl`. The color scheme of `SQL2Excel` is shown below.

<p align="center">
<img src="data/sql2excel_color_theme.png" style="display: block; margin: auto; width: 90%; height:200px;">
</p>

The default `Openpyxl` color theme is shown below.

<p align="center">
<img src="data/openpyxl_color_theme.png" style="display: block; margin: auto; width: 90%; height:200px;">
</p>

You can provide your own colors by overriding `PRIMARY_COLORS` in `sql2excel/config.py`.
You can do that programmatically by creating an instance of `Config` class and overriding its `PRIMARY_COLORS` attribute.
You can also revert back to `Openpyxl` default color by setting `OPENPYXL_COLORS` to `True`.

You can also disable the default `SQL2Excel` color theme without changing the default in `sql2excel/config.py` by passing `openpyxl_color=True` to the `plot` method of the `Chart` class. An example:

```python
chart = BarChart()
chart.plot(df, ws,  openpyxl_color=True)
```

### Font and Shape Colors

You can provide the colors for chart title and the labels for the x-axis and y-axis by specifying the _color name_ or the _color code_ in hex format. If a color name is provided, it must be chosen as a `prstClr` color as in the `openpxyl` guide [(here)](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.drawing.colors.html) where all available `prstClr` colors are listed.

The color of the section heading (the cell above the data and chart) can only provided as a color code. You can use online color pickers to find the color code of your choice such as [this color picker](https://htmlcolorcodes.com/color-picker/).

Shape colors (line colors and fill colors) should be provided as _color code_ in the hex format

# Developer Guide

## Chart Class

To complete ...

## Configuration

To complete ...

## Databases

To complete ...

## Tests

To complete ...

## Configuration

To complete ...

[Back to Top](#table-of-contents)

# Running Tests

As the package development is still in progress, you should run the tests (at least once when you first clone the git repo). To run the test follow the steps below.

- Open a new terminal, navigate to the project root (the directory that contains `README.md` file)
- Activate the virtual environment (only if it is NOT already active).

```bash
source .env/bin/activate
```

- Export the path to enable package discovery

```bash
export PYTHONPATH="$PYTHONPATH:$(pwd)";
export PYTHONPATH="$PYTHONPATH:$(pwd)/sql2excel";
```

- Run the tests

```bash
pytest tests/
```

[Back to Top](#table-of-contents)

# Contribution

While `SQL2Excel` is still in its early stages, it has potential to become a powerful tool for automating result reporting and data visualization. The contributions from the community are highly encouraged.

If you would like to contribute, whether by adding new features, improving the codebase, bug fixing, or providing feedback, your input is greatly appreciated. **Please feel free to open issues, submit pull requests, or suggest improvements.** If you would like to become a collaborator, please drop me an email at <a href="mailto:ahmedhassan@aims.ac.za">ahmedhassan@aims.ac.za</a>.

**Please try to unit test your code before submitting a pull request**. Testing `SQL2Excel` can be complicated in some cases (e.g. testing graphical properties of the chart). You can verify such features through comprehensive visual inspection.
Although the test coverage is not 100%, the core components of `SQL2Excel` are unit-tested.

[Back to Top](#table-of-contents)

# TO-DO

For features awaiting implementation, please see [TODO](TODO.md).

[Back to Top](#table-of-contents)
