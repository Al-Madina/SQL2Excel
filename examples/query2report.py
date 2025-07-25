"""
An example to show how to use sql2excel to generate reports if you work with
SQL from Python

@NOTE sqlalchemy 2.0 does not support positional parameters: 
https://github.com/sqlalchemy/sqlalchemy/issues/5178
Despite this, sql2excel implemented a crude way to handle positional parameters
as they are easier to work with. See the examples below.

@NOTE the queries below used the DVD rental database set up in a PostgreSQL: 
https://www.postgresqltutorial.com/postgresql-getting-started/postgresql-sample-database/
"""

import os

from db_params import connection_string

from sql2excel.report import Report
from sql2excel.sqlexec import QueryConfig

# Create a list (or any other sequence) that will hold all your QueryConfig objects
query_configs = []

# Query with no parameters
query = """SELECT rental_date::date, COUNT(*) AS rental_count
    FROM rental
    GROUP BY 1
    ORDER BY 1;"""
query = QueryConfig(sql=query, chart="line", section_heading="Query without parameters")
query_configs.append(query)


# Query with named parameter of type datetime
query = """SELECT rental_date::date, COUNT(*) AS rental_count
    FROM rental
    WHERE rental_date > :rental_date
    GROUP BY 1
    ORDER BY 1;"""
query = QueryConfig(
    sql=query,
    sql_params={"rental_date": "2005-07-31"},
    chart="bar",
    section_heading="Query with named parameter of type datetime",
)
query_configs.append(query)


# Positional parameter of type datetime
query = """SELECT rental_date::date, COUNT(*) AS rental_count
    FROM rental
    WHERE rental_date > ?
    GROUP BY 1
    ORDER BY 1;"""
# NOTE the comma at the end in `sql_params`
query = QueryConfig(
    sql=query,
    sql_params=("2005-07-31",),
    chart="bar",
    section_heading="Positional parameter of type datetime",
)
query_configs.append(query)


# Named numeric parameter
query = """SELECT length AS film_length, COUNT(*) AS film_count
    FROM film
    where length between :min_length and :max_length
    GROUP BY length
    ORDER BY film_length;"""
query = QueryConfig(
    sql=query,
    sql_params={"min_length": 100, "max_length": 110},
    chart="scatter",
    xlabel="Film length",
    ylabel="Number of films",
    title="Scatter Chart",
    section_heading="Named numeric parameter",
)
query_configs.append(query)


# positional numeric parameter
query = """SELECT length AS film_length, COUNT(*) AS film_count
    FROM film
    where length between ? and ?
    GROUP BY length
    ORDER BY film_length;"""

query = QueryConfig(
    sql=query,
    sql_params=(100, 110),
    chart="chart",
    section_heading="positional numeric parameter",
)
# queries_config.append(query)


# Named text parameter
query = """SELECT c.first_name, c.last_name, SUM(p.amount) AS total_amount
    FROM customer c
    JOIN payment p ON c.customer_id = p.customer_id
    WHERE c.last_name = :last_name
    GROUP BY c.first_name, c.last_name;"""
query = QueryConfig(
    sql=query,
    sql_params={"last_name": "Smith"},
    chart="chart",
    section_heading="Named parameter of type string",
)
query_configs.append(query)


# positional text parameter
query = """SELECT c.first_name, c.last_name, SUM(p.amount) AS total_amount
    FROM customer c
    JOIN payment p ON c.customer_id = p.customer_id
    WHERE c.last_name = ?
    GROUP BY c.first_name, c.last_name;"""
query = QueryConfig(
    sql=query,
    sql_params=("Smith",),
    chart="chart",
    section_heading="Positional parameter of type string",
)
query_configs.append(query)


# Yet more named text parameter
query = """SELECT c.first_name, c.last_name, SUM(p.amount) AS total_amount
    FROM customer c
    JOIN payment p ON c.customer_id = p.customer_id
    WHERE c.last_name like :name_part
    GROUP BY c.first_name, c.last_name;"""
query = QueryConfig(
    sql=query,
    sql_params={"name_part": "%St%"},
    chart="chart",
    section_heading="Named parameter of type string with wild card",
)
query_configs.append(query)


# Yet more positional text parameter
query = """SELECT c.first_name, c.last_name, SUM(p.amount) AS total_amount
    FROM customer c
    JOIN payment p ON c.customer_id = p.customer_id
    WHERE c.last_name like ?
    GROUP BY c.first_name, c.last_name;"""
query = QueryConfig(
    sql=query,
    sql_params=("%St%",),
    chart="chart",
    section_heading="Positional parameter of type string with wild card",
)
query_configs.append(query)


# NOTE ~ and ~* operators in Postgres will not work as it is not database-agnostic
# Even more positional text parameter
query = """SELECT c.first_name, c.last_name, SUM(p.amount) AS total_amount
    FROM customer c
    JOIN payment p ON c.customer_id = p.customer_id
    WHERE lower(c.last_name) like ?
    GROUP BY c.first_name, c.last_name;"""
query = QueryConfig(sql=query, sql_params=("%jo%",), chart="chart")
query_configs.append(query)


# Named sequence parameter
query = """SELECT c.customer_id,
            c.first_name,
            c.last_name,
            c.first_name || ' ' || c.last_name as full_name,
            SUM(p.amount) AS total_amount
    FROM customer c
            JOIN payment p ON c.customer_id = p.customer_id
    WHERE c.customer_id in :customer_ids
    GROUP BY 1, 2, 3, 4
    ORDER BY 5 DESC;"""
query = QueryConfig(
    sql=query, sql_params={"customer_ids": tuple(range(1, 6))}, chart="chart"
)
query_configs.append(query)


# Positional sequence parameter
query = """SELECT c.customer_id,
            c.first_name,
            c.last_name,
            c.first_name || ' ' || c.last_name as full_name,
            SUM(p.amount) AS total_amount
    FROM customer c
            JOIN payment p ON c.customer_id = p.customer_id
    WHERE c.customer_id in ?
    GROUP BY 1, 2, 3, 4
    ORDER BY 5 DESC;"""
query = QueryConfig(
    sql=query,
    # Note the comma after the list.
    sql_params=([0, 1, 2, 3, 4, 5],),
    chart="chart",
    section_heading="Passing a sequence as a positional parameter",
)
query_configs.append(query)


# Pivoting data to convert from long to wide format
query = """SELECT
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
            c.name IN ?
        GROUP BY
            1,
            2
        ORDER BY
            1,
            2;"""

query_config = QueryConfig(
    sql=query,
    # NOTE: note the comma after the nested insdie tuple.
    # Alternatively, you can pass a list
    # sql_params=(("Family", "Action", "Sports", "Music"),),
    sql_params=[("Family", "Action", "Sports", "Music")],  # Passing a list instead
    chart="bar",
    section_heading="Query result with pivoting data",
    # Pivoting parameters
    index="month_year",
    columns="category_name",
    values="total_payment",
)

query_configs.append(query_config)


####################################
#  Generate Report
####################################

report = Report(connection_string=connection_string)

file_name = os.path.join(
    os.path.dirname(os.path.dirname(__file__)), "data", "query2report.xlsx"
)


report.generate(query_configs, file_name=file_name)
