# Developer Guide

## Contribution
If you would like to contribute, whether by adding new features, improving the codebase, bug fixing, or providing feedback, your input is greatly appreciated. Please feel free to open issues, submit pull requests, or suggest improvements. If you would like to become a **collaborator**, please drop me an email at [ahmedhassan@aims.ac.za](mailto:ahmedhassan@aims.ac.za).

## AI Agents
You are welcome to use AI code agent to contribute to the project. However, note the following:

!!! warning
    If you use AI agents (Claude, Codex, Opencode, Cursor, etc), please do __not__ commit and push your agent config and artifacts to Github repo (e.g `.agent`, `.opencode`, `CLAUDE.md`, `AGENTS.md`, skills, MCP, etc). This repo does not assume any specific AI agent and publishing these will pollute it!

!!! tip
    The current `.gitignore` ignore default agent-specific config and artifacts. If your agent files or directories are tracked by `git`, please add them to `.gitignore`.

## Running Tests

- Clone Github repo

- Activate the virtual environment

```bash
source .env/bin/activate
```

- Install the development optional dependencies in the editable mode

```bash
pip install -e ".[test,docs,postgres]"
```

Repalce `postgres` with your database dialect (see [Supported Database](./index.md#supported-database))


- Run the tests

```bash
pytest tests/
```

- Run the test with the coverage

```bash
pytest --cov=sql2excel
```

## Databases

`SQL2Excel` is database-agnostic by design. It connects to any database supported by [SQLAlchemy](https://www.sqlalchemy.org/) via a standard connection URL. It is currently tested against PostgreSQL and MySQL. If you use another database and encounter issues, please open an issue on GitHub.

### Using Docker for Trial and Testing

Docker lets you spin up temporary database instances without installing them on your host machine. `SQL2Excel` is developed in machine hosting PostgreSQL (Docker is not used). However for MySQL, Docker was used and the instructions to set it up is below.  We will use the DvD rental (also known as Sakila) database to test `SQL2Excel` for MySQL.

#### PostgreSQL

Docker postgres guide is coming up ...


#### MySQL


**1. Download and Extract the Sakila Database**

First, download the database schema and data files directly from MySQL and extract them to your local machine. Run the following command (assuming Linux):

```bash
wget https://downloads.mysql.com/docs/sakila-db.tar.gz
tar -xvf sakila-db.tar.gz
```

This creates a folder named `sakila-db`.

**2. Start the MySQL Docker Container**

Pull and start the official MySQL image. We will use version 8.0.

```bash
docker run --name mysql-sakila \ 
    -e MYSQL_ROOT_PASSWORD=secretpassword \
    -p 3306:3306 \
    -d \
    mysql:8.0
```

- `--name mysql-sakila`: Gives the container an easy-to-remember name.

- `-e MYSQL_ROOT_PASSWORD=...`: Sets the administrative password.

- `-p 3306:3306`: Maps the default MySQL port from inside the container to your local machine, allowing `SQL2Excel` to connect to MySQL at `localhost:3306`.

- `-d`: Runs the container in the background (detached mode).

**3. Copy the Database Files into the Container**

To import the SQL files to MySQL, they need to be accessible from inside the running Docker container. You can copy the extracted folder directly into the container's root directory (or any other directory you prefer).

```bash
docker cp sakila-db mysql-sakila:/sakila-db
```

**4. Open the MySQL CLI**

Now, open an interactive terminal session inside the container and launch the MySQL command-line client to query the database.

```bash
docker exec -it mysql-sakila mysql -u root -p
```

When prompted for the password, enter your password you chose in Step 2 above. You should see `mysql>` in the terminal.

**5. Import the Schema and Data to MySQL**

From within the `mysql>` CLI, you can execute the SQL files you copied over. You must run the schema file first to create the tables, followed by the data file to populate them:

```sql
SOURCE /sakila-db/sakila-schema.sql;
SOURCE /sakila-db/sakila-data.sql;
```

**6. Your MySQL DB URL**

Your `db_url` that you will use to connect `SQL2Excel` to MySQL would be

```bash
"mysql+pymysql://root:<your-password>@localhost:3306/sakila"
```

Replace `<password>` with your password from Step 2.

That's it!

## Project Architecture

Understanding the project structure is essential before contributing.

### Core Modules

| Module | Description |
|--------|-------------|
| `sql2excel/report.py` | Top-level orchestrator that executes queries and writes Excel |
| `sql2excel/sqlexec.py` | SQL execution engine and query configuration |
| `sql2excel/parser.py` | SQL file parser for annotated queries |
| `sql2excel/chart.py` | All chart implementations (Line, Bar, Pie, etc.) |
| `sql2excel/config.py` | Default configuration constants |
| `sql2excel/excel_helper.py` | openpyxl utility layer |

### Chart Class Hierarchy

All chart types inherit from the `Chart` base class:

```
Chart
├── LineBasedChart (abstract)
│   ├── LineChart
│   └── RadarChart
├── TwoAxesChart (abstract)
│   ├── BarLineChart
├── BarChart
├── PieChart
├── AreaChart
├── ScatterChart
├── BubbleChart
├── StackedBarChart
├── SingleAxisBarLineChart
└── ImageChart
```

## Code Style

### Python Conventions

- Follow PEP 8 for code formatting
- Use type hints for all function signatures
- Write docstrings for public classes and methods
- Keep functions focused

### Naming Conventions

- **Classes**: PascalCase (`BarChart`, `QueryConfig`)
- **Functions**: snake_case (`parse_sql_file`, `generate_report`)
- **Constants**: UPPER_SNAKE_CASE (`DEFAULT_FONT_SIZE`)
- **Private methods**: Leading underscore (`_apply_style`)

## Adding a New Chart Type

To add a new chart type to `SQL2Excel`:

### Step 1: Create the Chart Class

In `sql2excel/chart.py`, add a new class inheriting from `Chart`:

```python
class YourNewChart(Chart):
    """Description of your chart type."""

    def __init__(self, config=None, excel_helper=None):
        super().__init__(config, excel_helper)

    def plot(self, df, worksheet, param1='x', param2='y'):
        """Specific description here

        Parameters
        ----------
        df : pd.DataFrame
            Data to plot.
        worksheet : openpyxl.worksheet.worksheet.Worksheet
            Target worksheet.
        param1 : int
            First keyword argument.
        param2 : int
            Second keyword argument
        """

        # Implementation here
        
        # Call the super method to finalize the plot
        # Note this is the private `_plot` (not the public `plot`)
        super()._plot(df, ws, **kwargs)
```

### Step 2: Register the Chart

Add the new class to `sql2excel/__init__.py`:

```python
from sql2excel.chart import YourNewChart

__all__ = [
    # ... existing exports
    "YourNewChart",
]
```

### Step 3: Add to Report

In `sql2excel/report.py`, update the chart type mapping:

```python
CHART_TYPES = {
    # ... existing types
    "your_new": YourNewChart,
}
```

### Step 4: Write Tests

Create tests in `tests/test_chart.py`:

```python
def test_your_new_chart(worksheet, sample_df):
    """Test YourNewChart plotting."""
    
    # Your assertions here 
```

## Testing Guidelines

- Use pytest
- Place tests in `tests/` directory
- Name test files `test_<module>.py`
- Use pytest fixtures defined in `tests/setup.py`


## Documentation

### Docstring Format

Use NumPy-style docstrings:

```python
def your_function(param1, param2):
    """Short summary.

    Longer description if needed.

    Parameters
    ----------
    param1 : str
        Description of param1.
    param2 : int
        Description of param2.

    Returns
    -------
    bool
        Description of return value.

    Raises
    ------
    ValueError
        If param2 is negative.
    """
```

### Building Documentation

```bash
# Install docs dependencies
pip install -e ".[docs]"

# Build static site
mkdocs build

# Serve documentation locally
mkdocs serve
```

## Pull Request Process

### Before Submitting

1. Write tests for your code. Please do __not__ push code that is not test
2. Run the full test suite: `pytest tests/`
3. Ensure all tests pass
4. Update documentation if adding new features


## Issue Labels

- `bug`: Something isn't working as expected
- `enhancement`: New feature or request
- `documentation`: Improvements to docs

## Getting Help

- Open an issue on GitHub for bugs or feature requests
- Email: [ahmedhassan@aims.ac.za](mailto:ahmedhassan@aims.ac.za)
- Review existing documentation before asking questions

