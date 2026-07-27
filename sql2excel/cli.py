"""
Script to execute SQL scripts and export the results to an Excel file.

Usage:
    ```sql2excel <sql_file> --dialect <dialect> --host <host> --port <port> --user <user> --password <pass> --dbname <db> --output <out_file>```

Example
-------
```
sql2excel query.sql
    --dialect postgresql
    --host localhost
    --port 5432
    --user username
    --password secret
    --dbname my_db
    --output report.xlsx
```
"""

# TODO support the use of .dbpass for Windows

import argparse
import csv
import os
import platform
import stat
import sys
from pathlib import Path

from sqlalchemy import URL

from sql2excel.parser import parse_sql_file
from sql2excel.report import Report


def _get_dbpass_path():
    """Return the path to .dbpass in the user's home directory."""
    return Path.home() / ".dbpass"


def _is_dbpass_perm_ok(filepath):
    """
    Ensure .dbpass has secure permissions.
    POSIX: file must be regular, owned by the user, mode = 0600, and not a symlink.
    Windows: skip (recommend Windows Credential Manager or python-keyring for storage).
    """
    if platform.system() == "Windows":
        # Checking permissions in Windows is complicated! Skip for now!
        return False

    try:
        # Reject symlinks to avoid TOCTOU and unintended indirections
        if filepath.is_symlink():
            print(
                f"WARNING: {filepath} is a symbolic link; refusing to use for security.",
                file=sys.stderr,
            )
            return False

        st = os.stat(filepath)
        # Ensure regular file
        if not stat.S_ISREG(st.st_mode):
            print(
                f"WARNING: {filepath} is not a regular file; refusing to use.",
                file=sys.stderr,
            )
            return False

        # Ensure file is owned by current user
        current_uid = getattr(os, "getuid", lambda: None)()
        if (
            hasattr(st, "st_uid")
            and current_uid is not None
            and st.st_uid != current_uid
        ):
            print(
                f"WARNING: {filepath} is not owned by the current user; refusing to use.",
                file=sys.stderr,
            )
            return False

        mode = stat.filemode(st.st_mode)
        if mode != "-rw-------":
            print(
                f"WARNING: Permissions on {filepath} are {mode}. "
                f"Expected owner read/write only. Ignoring file for security.",
                file=sys.stderr,
            )
            return False

        # permissions of .dbpass are OK
        return True

    except FileNotFoundError:
        return False

    except Exception as e:
        print(
            f"Warning: Error checking permissions on {filepath}: {e}", file=sys.stderr
        )
        return False


def _parse_dbpass_line(line: str):
    """
    Parse a .dbpass line using CSV logic to handle colons within quotes.
    Expected format: dialect:host:port:user:password:dbname
    """

    line = line.strip()
    if not line or line.startswith("#"):
        return None

    # Use csv.reader to handle quoted fields containing the delimiter
    try:
        reader = csv.reader([line], delimiter=":", quotechar='"', skipinitialspace=True)
        parts = next(reader)
    except Exception:
        return None

    if len(parts) != 6:
        return None

    # Return fields (dialect, host, port, user, password, dbname)
    return tuple(p.strip() if i != 4 else p for i, p in enumerate(parts))


def _match_field(pattern: str, value: str) -> bool:
    """Simple wildcard match: '*' matches anything, otherwise exact string compare."""
    return pattern == "*" or pattern == value


def _specificity_score(entry) -> int:
    """
    Score an entry based on specificity (non-wildcards).
    Higher score = more specific. Used to break ties when multiple lines match.
    """
    return sum(1 for p in entry[:6] if p != "*")


def _find_password_in_dbpass(
    dialect: str,
    host: str,
    port,
    user: str,
    dbname: str,
) -> str | None:
    """
    Scan .dbpass for a matching entry.
    Format per line: dialect:host:port:user:password:dbname
    Use '*' for wildcards in any field.

    Matching: exact string equality (case-sensitive) or '*' wildcard.
    If multiple entries match, the most specific (fewest '*') wins.
    Returns the first most-specific match's password, or None if not found.
    """
    dbpass_path = _get_dbpass_path()

    if not dbpass_path.exists():
        return None

    if not _is_dbpass_perm_ok(dbpass_path):
        return None

    matches = []
    try:
        with open(dbpass_path, "r", encoding="utf-8", errors="replace") as f:
            for line in f:
                parsed_line = _parse_dbpass_line(line)
                if not parsed_line:
                    continue

                f_dialect, f_host, f_port, f_user, f_password, f_dbname = parsed_line

                if (
                    _match_field(f_dialect, dialect)
                    and _match_field(f_host, host)
                    and _match_field(f_port, str(port))
                    and _match_field(f_user, user)
                    and _match_field(f_dbname, dbname)
                ):
                    matches.append(parsed_line)

        if not matches:
            return None

        # Choose the most specific match (fewest wildcards)
        best_match = max(matches, key=_specificity_score)
        return best_match[4]  # password
    except Exception as e:
        print(f"Warning: Error reading {dbpass_path}: {e}", file=sys.stderr)
        return None


def main():

    parser = argparse.ArgumentParser(
        prog="sql2excel",
        description="A program to execute SQL scripts and export the results to Excel.",
    )

    # Positional argument
    parser.add_argument("sql_file", help="Path to the SQL file")

    env_dialect = os.getenv("DB_DIALECT")
    env_host = os.getenv("DB_HOST")
    env_port = os.getenv("DB_PORT")
    env_user = os.getenv("DB_USER")
    env_pass = os.getenv("DB_PASSWORD")
    env_db = os.getenv("DB_NAME")

    parser.add_argument(
        "--dialect",
        required=False,
        default=env_dialect,
        help="Database dialect (as in SQLAlchemy)",
    )
    parser.add_argument(
        "--host",
        required=False,
        default=env_host,
        help="Database host address",
    )
    parser.add_argument(
        "--port",
        type=int,
        required=False,
        default=env_port,
        help="Database port",
    )
    parser.add_argument(
        "--user", required=False, default=env_user, help="Database username"
    )
    parser.add_argument(
        "--password",
        required=False,
        default=env_pass,
        help="Database password",
    )
    parser.add_argument(
        "--dbname", required=False, default=env_db, help="Database name"
    )
    parser.add_argument("--output", required=True, help="Path to output Excel file")

    args = parser.parse_args()

    if args.host is None:
        args.host = "localhost"

    if not args.user:
        parser.error(
            "Database username must be provided via --user or DB_USER environment variable."
        )

    if not args.dbname:
        parser.error(
            "Database name must be provided via --dbname or DB_NAME environment variable."
        )

    if args.port is None:
        parser.error(
            "Database port must be provided via --port or DB_PORT environment variable."
        )

    if not isinstance(args.port, int):
        try:
            args.port = int(args.port)
        except (ValueError, TypeError):
            parser.error(f"Invalid integer value for port: '{args.port}'")

    # Check if input SQL file exists
    sql_file = args.sql_file
    if not os.path.isfile(sql_file):
        parser.error(f"The input file '{sql_file}' does not exist.")

    # Validate input file extension
    ext = os.path.splitext(sql_file)[1].lower()
    if ext != ".sql":
        parser.error(
            f"Expected a SQL file (.sql) for input but received a file with extension: '{ext}'"
        )

    # Check if output directory exists
    output_file = args.output
    output_dir = os.path.dirname(output_file)
    # Only check if a directory path was actually provided (not just a filename)
    if output_dir and not os.path.isdir(output_dir):
        parser.error(
            f"The directory '{output_dir}' does not exist. Please create it first."
        )

    # Validate output file extension
    ext = os.path.splitext(output_file)[1].lower()
    if ext != ".xlsx":
        parser.error(
            f"Expected an Excel file (.xlsx) for output but received a file with extension: '{ext}'"
        )

    # Handle .dbpass
    f_password = _find_password_in_dbpass(
        args.dialect, args.host, args.port, args.user, args.dbname
    )
    password = f_password or args.password

    # Database URL
    try:
        db_url = URL.create(
            args.dialect,
            username=args.user,
            password=password,
            host=args.host,
            port=args.port,
            database=args.dbname,
        )
    except Exception as e:
        parser.error(f"Failed to construct Database URL: {e}")

    try:
        queries_config = parse_sql_file(sql_file)

        report = Report(db_url=db_url)

        # Generate the report and save the results to the output file
        report.generate(queries_config, file_name=output_file)

        print(f"Successfully executed. Results are exported to {output_file}")

    except Exception as e:
        print(f"Error during execution: {e}", file=sys.stderr)
        raise


if __name__ == "__main__":
    main()
