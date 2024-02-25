# Introduction
`xlsql` is a simple command-line tool that converts an Excel `.xlsx` spreadsheet into a `sqlite3` database.

# Installation
To install this utility using `pip` you can simply `pip install xlsql`

# Usage
To view command help, you can always run `xlsql --help`.

```
Usage: xlsql [OPTIONS] SPREADSHEET

Convert an Excel spreadsheet into a SQLite database.

Args:
    spreadsheet (str): The path to the Excel spreadsheet.

Options:
    --column, -c:  A column (or columns) to extract. Can be specified multiple times.
    --database:    The name of the database to create. (default: database.db)
    --force:       Overwrite an existing database.
    --sheet, -s:   A sheet (or sheets) to extract. Can be specified multiple times.
    --verbose, -v: Show verbose output.
    --version, -V: Show the xlsql version number.

Examples:
    xlsql ~/Documents/Example.xlsx
    # Creates: ~/Documents/example.db with all data included in the database.

    xlsql ~/Documents/Example.xlsx --verbose
    # Creates: ~/Documents/example.db, displaying verbose output while running.

    xlsql ~/Documents/Example.xlsx --database /tmp/example.db
    # Creates /tmp/example.db with all data included from the Excel sheet.

    xlsql ~/Documents/Example.xlsx --database /tmp/example.db --force
    # Overwrites the existing db with fresh content from the sheet!

    xlsql example.xlsx -c name -c id -c address -s people
    # Only select the name, id, and address columns from the people sheet.
```

# Contributing
To contribute to this project, please fork the repo and make your changes there.  Submit a PR back to this repo for review.

Be sure to install the dev dependencies, such as `pre-commit` and `black`.
