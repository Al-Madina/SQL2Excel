<div>
<p align="center">
<img src="assets/SQL2Excel_logo.png" style="display: block; margin: auto; width: 250px; height:250px;">
</p>
<p align="center">A Python library for exporting SQL queries and/or Pandas DataFrames to Excel, with support for chart generation and data visualization.</p>
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
    To install `SQL2Excel` dependencies for a specific database (e.g. Postgres), see [Database Support](#supported-database).

## Story of `SQL2Excel`

`SQL2Excel` stems from a project at [CREST](https://www0.sun.ac.za/crest/) to profile the scientific performance of the African countries across 52 scientific disciplines using the [Web of Science](https://clarivate.com/products/scientific-and-academic-research/research-discovery-and-workflow-solutions/webofscience-platform/) database. This resulted in generating 1560 reports, each with several performance indicators and different visualizations. `SQL2Excel` made such job possible!

## What SQL2Excel is Not

`SQL2Excel` is not meant as a library for reading/writing into Excel and it is not suitable for fine customizations of Excel. _Its purpose is to bridge the gap between SQL and Excel to improve productivity, automate repetitive tasks, and simplify data export and chart creation._ 

- If you want more control over databases, use database drivers or `SQLAlchemy`.
- If you want more control over Excel, use `Openpyxl` instead.

## Why Use SQL2Excel

The main reasons why you should consider using `SQL2Excel`:

- Integration with SQL
    - `SQL2Excel` executes your SQL script and exports all results to Excel including chart generation.
    - `SQL2Excel` Python API allows you to use Python to work with your databases and provides control over data export and chart generation.
    - `SQL2Excel` offers support for exporting data from parameterized queries. This improves the efficiency of repetitive tasks.
  
- Support for Pandas DataFrames

    - `SQL2Excel` exports DataFrames to Excel. While `Pandas` supports exporting dataframe to Excel, `SQL2Excel` offers more than just writing the dataframe to Excel. SQL2Excel can generate a full report from different dataframes including different visualizations.

- `Matplotlib` and externally generated images
    - Inserting `Matplotlib` figures and externally generated images is straightforward in `SQL2Excel`. This is useful if you require advanced charts that cannot be generated in Excel.
    
- Simple API
    - `SQL2Excel` offers simple API to work with SQL, export data, and generate charts.


## Supported Database

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


