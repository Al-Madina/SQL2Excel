"""
An example to show how to generate a report from a SQL file directly.
The SQL file needs to be annotated as in './sql2report.sql'
"""

import os

# NOTE db_params is a python file storing the database connection parameters.
# This is not the best practice to store database credentials. It is just an illustrative example.
from db_params import db_url, mysql_db_url

from sql2excel.parser import parse_sql_file
from sql2excel.report import Report

# Parent dir
parent_dir = os.path.dirname(os.path.dirname(__file__))

# Provide the path to the SQL file
queries_config = parse_sql_file(os.path.join(parent_dir, "sql", "sql2report_1.sql"))

# Create a report object
report = Report(db_url=db_url)  # Replace db_url with mysql_db_url to use MySQL database

# Generate the report and specify the file to save the results
report.generate(
    queries_config,
    file_name=os.path.join(parent_dir, "data", "sql2report_1.xlsx"),
)
