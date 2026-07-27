# Getting Started

This guide walks you through the three main ways to use SQL2Excel. 
## CLI

The fastest way to generate use `SQL2Excel` is through the command line — no Python code required.

- Annotate your SQL script with special directives (e.g. `-- chart`, `-- chart=bar`, etc)

```sql
-- my_script.sql

-- chart
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

The first directive `-- chart` exports the result to an Excel worksheet without creating any chart. The second `-- chart=bar` exports the result but also generated a pie chart. The results will be written in the _same_ worksheet by default.

- Activate your environment (assuming you created `.env`). See [installation](./index.md#installation).
```bash
source .env/bin/activate
```

- Run the command `sql2excel` in your terminal:

```bash
sql2excel my_script.sql \
    --dialect postgresql \
    --host localhost \
    --port 5432 \
    --user username \
    --password secret \
    --dbname my_db \
    --output report.xlsx
```

That's it! The CLI parses your SQL file, executes every annotated query against the database, and writes all results and charts into `report.xlsx`. See [Command Line Interface](./sql-script.md) for further information.

The command above assumes PostgreSQL. Change the `--dialect` flag if you are using a different database. The `--dialect` flag accepts any valid SQLAlchemy dialect and can be in the form `dialect+driver` if you are using a specific database driver (e.g. `mysql+pymysql2`).

!!! note
    Queries __without__ directives are skipped (not executed). Sometimes, you need to execute a query without exporting the result (e.g. to create a _temporary_ table that will be used in subsequent queries). In this case, use the directive `-- exec`.


!!! tip
    Always include a __space__ after -- (e.g., -- chart). While PostgreSQL allows `--chart` without a space, databases like MySQL strictly require it. Keeping the space ensures your SQL works across all supported drivers.


!!! warning
    While you can, you should __not__ type your password in the terminal directly. The password will be visible in your history. Instead, you can store your database credentials in a specific file in your home directory `~/.dbpass` or use environment variables. See [Command Line Interface](./sql-script.md) for further information. Currently, the `~/.dbpass` works for Linux and Mac only.

## Query with Positional Parameters

When you execute queries from Python, use `QueryConfig` to pair a SQL query with its parameters (if required) and chart options. Positional parameters are marked with `?` and supplied as a tuple in order:

```python
from sql2excel.sqlexec import QueryConfig

query = """
    SELECT length AS film_length, COUNT(DISTINCT film_id) AS film_count
    FROM film
    WHERE length BETWEEN ? AND ?
    GROUP BY length
    ORDER BY film_length;
"""

query_config = QueryConfig(
    sql=query,
    sql_params=(90, 120),       # positional: ? is replaced by 90 then 120
    chart="bar", # Export data and create bar chart
)
```

Then pass it to a `Report` to execute and save:

```python
from sql2excel.report import Report

report = Report(db_url="postgresql://user:pass@localhost:5432/my_db")
report.generate(query_config, file_name="output.xlsx")
```

The `db_url` is the same as SQLAlchemy database URI (sometimes referred to as a _connection string_). The positional parameters are processed in the _same_ order they are provided.


!!! note
    A one-element tuple needs a trailing comma: `(90,)`.

## Query with Named Parameters

Named parameters use `:param_name` syntax in the SQL and are supplied as a dictionary. They are easier to read and less error-prone than positional parameters:

```python
from sql2excel.sqlexec import QueryConfig
from sql2excel.report import Report

query = """
    SELECT rental_date::date, COUNT(DISTINCT rental_id) AS rental_count
    FROM rental
    WHERE rental_date > :start_date
    GROUP BY 1
    ORDER BY 1;
"""

query_config = QueryConfig(
    sql=query,
    sql_params={"start_date": "2005-07-31"},
    chart="line",
    section_heading="Daily rentals after July 2005",
    xlabel="Date",
    ylabel="Rental Count",
)

report = Report(db_url="postgresql://user:pass@localhost:5432/my_db")
report.generate(query_config, file_name="rentals.xlsx")
```

Note the `:` in parameter name `:start_date` in SQL query. Multiple named parameters work the same way — just add more keys to the dictionary and corresponding placeholders in the SQL. For example, `{"param1": value1, "param2": value2, ...}` then in your SQL query use `:param1`, `:param2`, etc.


!!! note
    Multiple SQL queries can be exported by creating a list of `QueryConfig` objects and passing it to `Report.generate` as the first argument. See [Generating Report from QueryConfig](query.md/#generating-report-from-queryconfig).



## Pandas Dataframes

If your data already in a DataFrame, you can direclty export it to Excel and generate charts:

```python
import pandas as pd
import openpyxl as xl
from sql2excel.chart import BarChart

df = pd.DataFrame({
    "Region": ["North", "South", "East", "West"],
    "Revenue": [120000, 95000, 110000, 87000],
})

# Create an excel workbook
wb = xl.Workbook()

# Use the active worksheet
ws = wb.active

chart = BarChart()
# This will export the data and generate a bar chart
chart.plot(df, ws, title="Revenue by Region")

wb.save("revenue.xlsx")
```

For writing data __without__ a chart, use `Chart.write_dataframe` instead:

```python
from sql2excel.chart import Chart

wb = xl.Workbook()
ws = wb.active

chart = Chart()
chart.write_dataframe(df, ws)

wb.save("data_only.xlsx")
```

Any `Chart` subclass (`LineChart`, `PieChart`, `ScatterChart`, etc.) works the same way — create it, then call `plot` with your DataFrame and worksheet. See [Supported Charts](chart.md) for the full list.

Congratulations! Now, you can use `SQL2Excel` from the command-line interface and Python.
