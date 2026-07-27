<div>
<p align="center">
  <img
    src="https://raw.githubusercontent.com/Al-Madina/SQL2Excel/main/data/SQL2Excel_logo.png"
    width="250"
    height="250"
    alt="SQL2Excel logo">
</p>
<p align="center">A Python library for exporting SQL queries and/or Pandas DataFrames to Excel, with support for chart generation and data visualization.</p>


<p align="center">
  <!-- PyPI Version -->
  <a href="https://pypi.org/project/SQL2Excel/">
    <img src="https://img.shields.io/pypi/v/SQL2Excel?color=blue&logo=pypi&logoColor=white" alt="PyPI Version">
  </a>

  <!-- Python Supported Versions -->
  <a href="https://pypi.org/project/SQL2Excel/">
    <img src="https://img.shields.io/pypi/pyversions/SQL2Excel?color=red&logo=python&logoColor=white" alt="Python Version">
  </a>

  <!-- CI Tests Workflow Status (Dynamic via ci.yml) -->
  <a href="https://github.com/Al-Madina/SQL2Excel/actions/workflows/ci.yml">
    <img src="https://img.shields.io/github/actions/workflow/status/Al-Madina/SQL2Excel/ci.yml?branch=main&label=tests&color=blue&logo=github" alt="CI Build Status">
  </a>

  <!-- Code Coverage Badge (Dynamic via Codecov) -->
  <a href="https://codecov.io/gh/Al-Madina/SQL2Excel">
    <img src="https://codecov.io/gh/Al-Madina/SQL2Excel/branch/main/graph/badge.svg" alt="Code Coverage">
  </a>

  <!-- License -->
  <a href="https://github.com/Al-Madina/SQL2Excel/blob/main/LICENSE">
    <img src="https://img.shields.io/github/license/Al-Madina/SQL2Excel?color=green" alt="License">
  </a>

  <!-- PyPI Downloads -->
  <a href="https://pypistats.org/packages/sql2excel">
    <img src="https://img.shields.io/pypi/dm/SQL2Excel?color=orange" alt="PyPI Downloads">
  </a>

  <!-- Code style -->
  <img src="https://img.shields.io/badge/code%20style-black-000000.svg" alt="Code Style: Black">

  <!-- Documentation -->
  <img src="https://img.shields.io/badge/docs-online-blue?logo=readthedocs" alt="Documentation">

  <!-- Stars -->
  <img src="https://img.shields.io/github/stars/Al-Madina/SQL2Excel?style=social" alt="GitHub Stars">
</p>
</div>


# SQL2Excel

`SQL2Excel` exports result sets from SQL queries and/or Pandas DataFrames into Excel including chart generation and data visualization with a strong focus on __automating__  and __simplifying__ the process of exporting data from SQL/Python to Excel. It supports various chart types and allows for basic chart customization.

## Installation

- Create a virtual environment (`.env`)

```bash
python -m venv .env
```

- Activate the environment

```bash
source .env/bin/activate
```


- Install `SQL2Excel` and its dependencies

```bash
pip install SQL2Excel
```

- Verify the Installation

```bash
python -c "import sql2excel; print(f'version: {sql2excel.__version__}')";
```

!!! tip
    To install `SQL2Excel` dependencies for a specific database (e.g. Postgres), see [Database Support](#database-support).

## Getting Started

`SQL2Excel` requires you to add special comments (called _directives_) to your SQL script to instruct it how to export the query result.

```sql
-- my_script.sql

-- chart
SELECT year, revenue FROM annual_sales;

-- chart=bar
SELECT region, SUM(revenue) AS total_revenue
FROM regional_sales
GROUP BY region
ORDER BY total_revenue DESC;

-- exec
CREATE TEMPORARY TABLE tmp_table AS ...

-- This query will be SKIPPED
SELECT * from foo.bar
```

- `-- chart` directive will export the query result to Excel without generating any chart because the chart type is not provided.

- `-- char=bar` directive will export the query result to Excel and generate a bar chart from the result.

- `-- exec` executes the third query without exporting any result. This is useful for DDL, DML, and temporary tables that might be used in subsequent query.

- The fourth query will be ignored because it does not contain any directives.

Then, run the command `sql2excel` in your terminal:

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

That's it!



## Story of `SQL2Excel`

`SQL2Excel` stems from a project at [CREST](https://www0.sun.ac.za/crest/) to profile the scientific performance of the African countries across 52 scientific disciplines using the [Web of Science](https://clarivate.com/products/scientific-and-academic-research/research-discovery-and-workflow-solutions/webofscience-platform/) database. This resulted in generating 1560 reports, each with several performance indicators and different visualizations. `SQL2Excel` made such job possible!


## Why Use SQL2Excel

The main reasons why you should consider using `SQL2Excel` despite the existence of `Openpyxl` are:

- Integration with SQL
    - `SQL2Excel` executes your SQL script and exports all results to Excel including chart generation.
    - `SQL2Excel` offers support for exporting data from parameterized queries. This improves the efficiency of repetitive tasks.
  
- Support for Pandas DataFrames

    - `SQL2Excel` exports DataFrames to Excel. While `Pandas` supports exporting dataframe to Excel, `SQL2Excel` offers more than just writing the dataframe to Excel. SQL2Excel can generate a full report from different dataframes including different visualizations.

- `Matplotlib` and externally generated images
    - Inserting `Matplotlib` figures and externally generated images is straightforward in `SQL2Excel`. This is useful if you require advanced charts that cannot be generated in Excel.
- Simple API
    - `SQL2Excel` offers simple API for data export and chart generation.

`SQL2Excel` is not meant as a library for reading/writing into Excel and it is not suitable for fine customizations of Excel. _Its purpose is to improve productivity, automate repetitive tasks, and simplify data export and chart creation._
If you want more control over excel from within Python, use `Openpyxl` instead.

## Database Support

`SQL2Excel` works with any database supported by [SQLAlchemy](https://www.sqlalchemy.org/). This includes PostgreSQL, MySQL, Oracle, and SQL Server. It is currently tested against PostgreSQL and MySQL. 


### PostgreSQL

To install `SQL2Excel`for postgres you have two options:

- Using `pyscopg2-binary`

```bash
pip install SQL2Excel[postgres]
```

- Using pycopg3 

```bash
pip install SQL2Excel[postgres-psycopg]
```


### MySQL

```bash
pip install SQL2Excel[mysql]
```

### Oracle

```bash
pip install SQL2Excel[oracle]
```

### SQL Server

```bash
pip install SQL2Excel[mssql]
```

### Snowflake

```bash
pip install SQL2Excel[snowflake]
```

### Databricks 

```bash
pip install SQL2Excel[databricks]
```

### BigQuery

```bash
pip install SQL2Excel[bigquery]
```


## Official Documentation
The official documentation is available [here](https://Al-Madina.github.io/SQL2Excel/).