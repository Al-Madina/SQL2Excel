import argparse
import os
import stat
import sys
from unittest.mock import MagicMock, mock_open, patch

import pytest

from sql2excel.cli import (
    _find_password_in_dbpass,
    _is_dbpass_perm_ok,
    _match_field,
    _parse_dbpass_line,
    _specificity_score,
    main,
)


def test_parse_dbpass_line():
    # Valid line
    line = "postgres:localhost:5432:user:pass:db"
    expected = ("postgres", "localhost", "5432", "user", "pass", "db")
    assert _parse_dbpass_line(line) == expected

    # Line with quoted password containing colon
    line = 'postgres:localhost:5432:user:"pass:word":db'
    expected = ("postgres", "localhost", "5432", "user", "pass:word", "db")
    assert _parse_dbpass_line(line) == expected

    # Comment line
    assert _parse_dbpass_line("# This is a comment") is None

    # Empty line
    assert _parse_dbpass_line("") is None

    # Malformed line (too few fields)
    assert _parse_dbpass_line("postgres:localhost") is None


def test_match_field():
    assert _match_field("*", "anything") is True
    assert _match_field("value", "value") is True
    assert _match_field("value", "other") is False


def test_specificity_score():
    # All wildcards
    entry = ("*", "*", "*", "*", "pass", "*")
    # Score counts non-wildcards in the first 6 fields
    # "pass" is not "*", so score is 1
    assert _specificity_score(entry) == 1

    # Some specific fields
    entry = ("postgres", "localhost", "*", "*", "pass", "*")
    # "postgres", "localhost", "pass" are not "*" -> score 3
    assert _specificity_score(entry) == 3

    # All specific
    entry = ("postgres", "localhost", "5432", "user", "pass", "db")
    assert _specificity_score(entry) == 6


@patch("sql2excel.cli.platform.system")
def test_is_dbpass_perm_ok_windows(mock_system):
    mock_system.return_value = "Windows"
    assert _is_dbpass_perm_ok(MagicMock()) is False


@patch("sql2excel.cli.platform.system")
@patch("sql2excel.cli.os.stat")
@patch("sql2excel.cli.os.getuid", create=True)
def test_is_dbpass_perm_ok_posix(mock_getuid, mock_stat, mock_system):
    mock_system.return_value = "Linux"
    mock_path = MagicMock()

    # Case 1: Symlink
    mock_path.is_symlink.return_value = True
    assert _is_dbpass_perm_ok(mock_path) is False

    # Case 2: Not a regular file
    mock_path.is_symlink.return_value = False
    st = MagicMock()
    st.st_mode = stat.S_IFDIR
    mock_stat.return_value = st
    assert _is_dbpass_perm_ok(mock_path) is False

    # Case 3: Wrong owner
    st.st_mode = stat.S_IFREG | 0o600
    st.st_uid = 1000
    mock_getuid.return_value = 1001  # Different UID
    mock_stat.return_value = st
    assert _is_dbpass_perm_ok(mock_path) is False

    # Case 4: Wrong permissions (readable by group/others)
    mock_getuid.return_value = 1000
    st.st_uid = 1000
    st.st_mode = stat.S_IFREG | 0o644
    mock_stat.return_value = st
    assert _is_dbpass_perm_ok(mock_path) is False

    # Case 5: OK
    st.st_mode = stat.S_IFREG | 0o600
    mock_stat.return_value = st
    assert _is_dbpass_perm_ok(mock_path) is True

    # Case 6: FileNotFoundError
    mock_stat.side_effect = FileNotFoundError
    assert _is_dbpass_perm_ok(mock_path) is False


@patch("sql2excel.cli._get_dbpass_path")
@patch("sql2excel.cli._is_dbpass_perm_ok")
def test_find_password_in_dbpass(mock_perm_ok, mock_get_path):
    mock_path = MagicMock()
    mock_get_path.return_value = mock_path

    # File does not exist
    mock_path.exists.return_value = False
    assert _find_password_in_dbpass("pg", "loc", 5432, "u", "db") is None

    # File exists but perms bad
    mock_path.exists.return_value = True
    mock_perm_ok.return_value = False
    assert _find_password_in_dbpass("pg", "loc", 5432, "u", "db") is None

    # File ok, read content
    mock_perm_ok.return_value = True
    content = """
    # Comment
    postgres:localhost:5432:user:pass1:db
    postgres:*:*:*:pass2:*
    """
    with patch("builtins.open", mock_open(read_data=content)):
        # Exact match (most specific)
        pw = _find_password_in_dbpass("postgres", "localhost", 5432, "user", "db")
        assert pw == "pass1"

        # Wildcard match
        pw = _find_password_in_dbpass("postgres", "other", 5432, "other", "other")
        assert pw == "pass2"

        # No match
        pw = _find_password_in_dbpass("mysql", "loc", 3306, "root", "db")
        assert pw is None


@patch("sql2excel.cli.argparse.ArgumentParser")
@patch("sql2excel.cli.os.path.isfile")
@patch("sql2excel.cli.os.path.isdir")
@patch("sql2excel.cli.parse_sql_file")
@patch("sql2excel.cli.Report")
@patch("sql2excel.cli.URL.create")
def test_main_success(
    mock_url_create,
    mock_report_cls,
    mock_parse,
    mock_isdir,
    mock_isfile,
    mock_parser_cls,
):
    # Setup mocks
    mock_parser = mock_parser_cls.return_value
    mock_parser.parse_args.return_value = argparse.Namespace(
        sql_file="query.sql",
        dialect="postgres",
        host="localhost",
        port=5432,
        user="user",
        password="password",
        dbname="db",
        output="output.xlsx",
    )
    mock_isfile.return_value = True
    mock_isdir.return_value = True

    mock_report_instance = mock_report_cls.return_value

    # Run main
    main()

    # Assertions
    mock_parse.assert_called_once_with("query.sql")
    mock_report_cls.assert_called_once()
    mock_report_instance.generate.assert_called_once()


@patch("sql2excel.cli.argparse.ArgumentParser")
@patch("sql2excel.cli.os.path.isfile")
@patch("sql2excel.cli.os.path.isdir")
@patch("sql2excel.cli.parse_sql_file")
@patch("sql2excel.cli.Report")
@patch("sql2excel.cli.URL.create")
def test_main_bare_output_filename_skips_dir_check(
    mock_url_create,
    mock_report_cls,
    mock_parse,
    mock_isdir,
    mock_isfile,
    mock_parser_cls,
):
    # `--output report.xlsx` (as in the README) has no directory part, so the
    # output-directory check must not run: os.path.isdir("") is False.
    mock_parser = mock_parser_cls.return_value
    mock_parser.error.side_effect = SystemExit
    mock_parser.parse_args.return_value = argparse.Namespace(
        sql_file="query.sql",
        dialect="postgres",
        host="localhost",
        port=5432,
        user="user",
        password="password",
        dbname="db",
        output="report.xlsx",
    )
    mock_isfile.return_value = True
    mock_isdir.return_value = False

    main()

    mock_parser.error.assert_not_called()
    mock_report_cls.return_value.generate.assert_called_once()


@patch("sql2excel.cli.argparse.ArgumentParser")
@patch("sql2excel.cli.os.path.isfile")
def test_main_input_file_not_found(mock_isfile, mock_parser_cls):
    mock_parser = mock_parser_cls.return_value
    mock_parser.error.side_effect = SystemExit

    mock_parser.parse_args.return_value = argparse.Namespace(
        sql_file="missing.sql",
        dialect="postgresql",
        host="host",
        port=5432,
        user="user",
        password="password",
        dbname="db",
        output="out.xlsx",
    )
    mock_isfile.return_value = False

    with pytest.raises(SystemExit):
        main()

    mock_parser.error.assert_called_with("The input file 'missing.sql' does not exist.")


@patch("sql2excel.cli.argparse.ArgumentParser")
@patch("sql2excel.cli.os.path.isfile")
def test_main_invalid_sql_extension(mock_isfile, mock_parser_cls):
    mock_parser = mock_parser_cls.return_value
    mock_parser.error.side_effect = SystemExit

    mock_parser.parse_args.return_value = argparse.Namespace(
        sql_file="query.txt",
        dialect="postgresql",
        host="host",
        port=5432,
        user="user",
        password="password",
        dbname="db",
        output="out.xlsx",
    )
    mock_isfile.return_value = True

    with pytest.raises(SystemExit):
        main()

    mock_parser.error.assert_called()
    assert "Expected a SQL file" in mock_parser.error.call_args[0][0]


@patch("sql2excel.cli.argparse.ArgumentParser")
@patch("sql2excel.cli.os.path.isfile")
@patch("sql2excel.cli.os.path.isdir")
def test_main_invalid_output_dir(mock_isdir, mock_isfile, mock_parser_cls):
    mock_parser = mock_parser_cls.return_value
    mock_parser.error.side_effect = SystemExit

    mock_parser.parse_args.return_value = argparse.Namespace(
        sql_file="query.sql",
        dialect="postgresql",
        host="host",
        port=5432,
        user="user",
        password="password",
        dbname="db",
        output="/nonexistent/out.xlsx",
    )
    mock_isfile.return_value = True
    mock_isdir.return_value = False

    with pytest.raises(SystemExit):
        main()

    mock_parser.error.assert_called()
    assert (
        "directory '/nonexistent' does not exist" in mock_parser.error.call_args[0][0]
    )


@patch("sql2excel.cli.argparse.ArgumentParser")
@patch("sql2excel.cli.os.path.isfile")
@patch("sql2excel.cli.os.path.isdir")
def test_main_invalid_output_extension(mock_isdir, mock_isfile, mock_parser_cls):
    mock_parser = mock_parser_cls.return_value
    mock_parser.error.side_effect = SystemExit

    mock_parser.parse_args.return_value = argparse.Namespace(
        sql_file="query.sql",
        dialect="postgresql",
        host="host",
        port=5432,
        user="user",
        password="password",
        dbname="db",
        output="out.csv",
    )
    mock_isfile.return_value = True
    mock_isdir.return_value = True

    with pytest.raises(SystemExit):
        main()

    mock_parser.error.assert_called()
    assert "Expected an Excel file" in mock_parser.error.call_args[0][0]


import pytest


@patch("sql2excel.cli.argparse.ArgumentParser")
@patch("sql2excel.cli.os.path.isfile")
@patch("sql2excel.cli.os.path.isdir")
@patch("sql2excel.cli.parse_sql_file")
@patch("sql2excel.cli.Report")
@patch("sql2excel.cli.URL.create")
def test_main_execution_error(
    mock_url_create,
    mock_report_cls,
    mock_parse,
    mock_isdir,
    mock_isfile,
    mock_parser_cls,
):
    mock_parser = mock_parser_cls.return_value
    mock_parser.parse_args.return_value = argparse.Namespace(
        sql_file="query.sql",
        dialect="postgres",
        host="localhost",
        port=5432,
        user="user",
        password="password",
        dbname="db",
        output="output.xlsx",
    )
    mock_isfile.return_value = True
    mock_isdir.return_value = True

    # Simulate error during parsing
    mock_parse.side_effect = Exception("Parsing error")

    # Verify that calling main() raises the Exception
    with pytest.raises(Exception, match="Parsing error"):
        main()


@patch("sql2excel.cli.argparse.ArgumentParser")
@patch("sql2excel.cli.os.path.isfile")
@patch("sql2excel.cli.os.path.isdir")
@patch("sql2excel.cli.parse_sql_file")
@patch("sql2excel.cli.Report")
@patch("sql2excel.cli.URL.create")
@patch("sql2excel.cli._find_password_in_dbpass")
def test_main_default_host(
    mock_find_pass,
    mock_url_create,
    mock_report_cls,
    mock_parse,
    mock_isdir,
    mock_isfile,
    mock_parser_cls,
):
    # Setup mocks
    mock_find_pass.return_value = None
    mock_parser = mock_parser_cls.return_value
    args = argparse.Namespace(
        sql_file="query.sql",
        dialect="postgres",
        host=None,  # Host is None
        port=5432,
        user="user",
        password="password",
        dbname="db",
        output="output.xlsx",
    )
    mock_parser.parse_args.return_value = args
    mock_isfile.return_value = True
    mock_isdir.return_value = True

    # Run main
    main()

    # Assert that args.host was set to 'localhost'
    assert args.host == "localhost"

    # Check that URL.create was called with the correct host
    mock_url_create.assert_called_once_with(
        "postgres",
        username="user",
        password="password",
        host="localhost",  # Check this
        port=5432,
        database="db",
    )


@patch("sql2excel.cli.argparse.ArgumentParser")
@pytest.mark.parametrize(
    "missing_arg, error_message",
    [
        (
            "user",
            "Database username must be provided via --user or DB_USER environment variable.",
        ),
        (
            "dbname",
            "Database name must be provided via --dbname or DB_NAME environment variable.",
        ),
        (
            "port",
            "Database port must be provided via --port or DB_PORT environment variable.",
        ),
    ],
)
def test_main_missing_required_db_args(mock_parser_cls, missing_arg, error_message):
    mock_parser = mock_parser_cls.return_value
    mock_parser.error.side_effect = SystemExit

    args = {"user": "user", "dbname": "db", "port": 5432, "host": "somehost"}
    args[missing_arg] = None

    mock_parser.parse_args.return_value = argparse.Namespace(**args)

    with pytest.raises(SystemExit):
        main()

    mock_parser.error.assert_called_with(error_message)


@patch("sql2excel.cli.argparse.ArgumentParser")
def test_main_invalid_port_type(mock_parser_cls):
    mock_parser = mock_parser_cls.return_value
    mock_parser.error.side_effect = SystemExit

    mock_parser.parse_args.return_value = argparse.Namespace(
        user="user", dbname="db", port="not-an-int", host="somehost"
    )

    with pytest.raises(SystemExit):
        main()

    mock_parser.error.assert_called_with("Invalid integer value for port: 'not-an-int'")
