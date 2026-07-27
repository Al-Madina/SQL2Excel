# Generate Excel from SQL Queries Executed Within Python

If you work with SQL from **within Python**, you do **not** need to create a `.sql` file to use `SQL2Excel`. Instead, you can export the results and generate charts using `SQL2Excel` Python API. For this purpose you need to create `QueryConfig` object directly from your SQL query, then create a `Report` object and invoke its `generate` method (see below).

## Basic Example


```python
from sql2excel.sqlexec import QueryConfig
from sql2excel.report import Report

# Your SQL Query
query = """SELECT rental_date::date, COUNT(DISTINCT rental_id) AS rental_count
           FROM rental
           GROUP BY 1
           ORDER BY 1;"""

# Create a QueryConfig object. Note `chart="line"`
query = QueryConfig(sql=query, chart="line")

# Create a report object
report = Report(db_url="postgresql://user:pass@localhost:5432/my_db")

# Generate the report and save it to an Excel file
report.generate(
    query_config, fname="path/to/excel/file.xlsx"
)
```

That's it! `SQL2Excel` offers support for paramterized queries that can make repetitive tasks more efficient. Read below to learn more.


## QueryConfig

`QueryConfig` defines everything `SQL2Excel` needs to run your query and export data to Excel. To construct a `QueryConfig` object you need to provide the following arguments:

- SQL Query (`sql`): This is the SQL query you want to run.

- SQL Parameters (`sql_params`): If your query needs some parameters to run.

- Excel Options: _keyword arguments_ to control data export and visualization in Excel.


## Passing Named Parameters to SQL Queries

`SQL2Excel` uses `SQLAlchemy` to connect and query databases. Therefore, you can pass parameters to
your query the same way you would if you used `SQLAlchemy`. Named parameters is just a Python dictionary with the key representing the parameter name and the value representing the parameter value. In the query below `rental_date_param` is a
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
    # Named parameters defined as a dict
    sql_params={"rental_date_param": "2005-07-31"},
    chart="bar",
)
```

You can pass multiple parameters `{"param1": value1, "param2": value2, ...}` and in you SQL prefix your parameters with `:` (e.g., `:param1`, `:param2`, etc).

## Passing Positional Parameters to SQL Queries

You can pass positional parameters to your SQL queries by marking them with `?` and providing the values of the parameters as a `Sequence` (e.g. list, tuple, etc). Note that _positional parameters are processed in the same order they are provided_.

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

!!! tip
    If you want to pass a one-element tuple, remember to include the comma at the end as `(my_param_value,)`.

!!! note
    SQLAlchemy 2.0 does not support positional parameters [(here)](https://github.com/sqlalchemy/sqlalchemy/issues/5178). The issue cited above suggests using `exec_driver_sql`. However, passing parameters this way is not DB-agnostic and will require manual adjustment. `SQL2Excel` uses an alternative approach, which is database-agnostic, to handle positional parameters.

## Data and Chart Options

As mentioned, any options related to Excel should be passed as _keyword arguments_. An example
of chart customization is shown below.

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

!!! tip
    To write data without generating chart, use the option `chart="chart"`.

For a full list customization options, see [Featured Options](./miscellaneous.md#featured-options).
## Generating Report from QueryConfig

Once you define your `QueryConfig` object, you can export the result set of `QueryConfig` by creating a `Report` object and call its `generate` method as in the example below.

```python
from db_params import db_url

# An example of connection string:
# 'postgresql://username:password@localhost:5432/database_name'

query_config = .... # Create a QueryConfig object as explained above

report = Report(db_url=db_url)

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