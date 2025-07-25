import os
from setuptools import setup, find_packages


def version():
    with open(
        os.path.join(os.path.dirname(__file__), "sql2excel", "__init__.py"), "r"
    ) as file:
        for line in file:
            if "__version__" in line:
                version = line.split("=")[1].strip().strip('"')
                return version


def install_requires():
    with open(os.path.join(os.path.dirname(__file__), "requirements.txt"), "r") as file:
        install_requires = [
            line.strip() for line in file if line.strip() and not line.startswith("#")
        ]

    return install_requires


def readme():
    return open(os.path.join(os.path.dirname(__file__), "README.md"), "r").read()


setup(
    name="SQL2Excel",
    version=version(),  # still beta version
    license="MIT",
    description="""Exporting SQL queries and/or Pandas DataFrames to Excel, with support for chart generation and data visualization.""",
    long_description=readme(),
    long_description_content_type="text/markdown",
    home_page="",
    url="",
    author="Ahmed Hassan",
    author_email="ahmedhassan@aims.ac.za",
    maintainer="Ahmed Hassan",
    maintainer_email="ahmedhassan@aims.ac.za",
    packages=find_packages(),
    install_requires=install_requires(),
)
