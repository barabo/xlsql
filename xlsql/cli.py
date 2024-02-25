"""xlsql converts an Excel .xlsx spreadsheet into a sqlite3 database.

xlsql converts an Excel .xlsx spreadsheet into a sqlite3 database.

Usage: xlsql [OPTIONS] SPREADSHEET

Convert an Excel spreadsheet into a SQLite database.

Args:
    spreadsheet (str): The path to the Excel spreadsheet.

Options:
    --column, -c: A column (or columns) to extract. Can be specified multiple times.
    --database: The name of the database to create. (default: database.db)
    --force: Overwrite an existing database.
    --sheet, -s: A sheet (or sheets) to extract. Can be specified multiple times.
    --verbose, -v: Show verbose output.

Examples:
    xlsql ~/Documents/Example.xlsx  # Creates: ~/Documents/example.db
    xlsql ~/Documents/Example.xlsx --database /tmp/example.db
    xlsql ~/Documents/Example.xlsx --database /tmp/example.db --force  # Overwrites existing db!

"""
import click
import openpyxl
import sqlite3
from pathlib import Path


def normalize(name: str) -> str:
    """
    Normalize a given name by converting it to lowercase, removing
    non-printable characters, replacing hyphens and spaces with underscores.

    Args:
        name (str): The name to be normalized.

    Returns:
        str: The normalized name.
    """
    name = "EMPTY" if name is None else name
    name = name.lower().strip()
    name = "".join(c for c in name if c.isprintable())
    return name.replace("-", "_").replace(" ", "_")


def get_column_names(sheet_name: str, headings: list[str], log: any) -> list[str]:
    """
    Get unique column names for a given sheet.

    Args:
        sheet_name (str): The name of the sheet.
        headings (list[str]): The list of headings.
        log (any): The logging function.

    Returns:
        list[str]: The list of distinct, normalized column names.
    """
    seen = {}
    column_names = []
    for heading in headings:
        normalized = normalize(heading)
        while normalized in seen:
            suffix = seen[normalized]
            seen[normalized] += 1
            normalized = f"{normalized}_{suffix}"
            log(
                f"WARN: duplicate heading in {sheet_name}[{heading}]: renaming to: {normalized}"
            )
        else:
            seen[normalized] = 2
        column_names.append(normalized)
    assert len(column_names) == len(set(column_names))
    return column_names


@click.command()
@click.argument(
    "spreadsheet", type=click.Path(exists=True, readable=True, dir_okay=False)
)
@click.option(
    "--column",
    "-c",
    multiple=True,
    help="A column (or columns) to extract. Can be specified multiple times.",
)
@click.option(
    "--database",
    type=click.Path(exists=False),
    default="database.db",
    help="The name of the database to create.",
)
@click.option(
    "--force",
    is_flag=True,
    type=bool,
    default=False,
    help="Overwrite an existing database.",
)
@click.option(
    "--sheet",
    "-s",
    multiple=True,
    help="A sheet (or sheets) to extract. Can be specified multiple times.",
)
@click.option(
    "--verbose",
    "-v",
    is_flag=True,
    type=bool,
    default=False,
    help="Show verbose output.",
)
def main(spreadsheet, column, database, force, sheet, verbose) -> None:
    """
    Convert an Excel spreadsheet into a SQLite database.

    Args:
        spreadsheet (str): The path to the Excel spreadsheet.
        column (list[str]): The name of a column or columns to extract.
        database (str): The name of the database to create.
        force (bool): Flag to overwrite an existing database.
        sheet (list[str]): The name of the sheet or sheets to extract.
        verbose (bool): Flag to show verbose output.

    Returns:
        None

    Raises:
        click.ClickException: If the destination database already exists and the force flag is not set.
    """

    def log(message: str) -> None:
        if verbose:
            print(message)

    # Ensure that the target database won't be overwritten, or that it's OK to
    # overwrite it.
    existing = Path(database)
    if database and existing.exists() and existing.stat().st_size:
        log(f"Destination database already exists: {database}")
        if force:
            log("Overwriting due to --force flag.")
            existing.unlink()
        else:
            raise click.ClickException(
                f"Cowardly refusing to overwrite existing db: {database}"
            )

    try:
        # Load the spreadsheet.
        log(f"Reading contents of speadsheet file: {spreadsheet}")
        workbook = openpyxl.load_workbook(spreadsheet)

        # Create a new SQLite database and connect to it.
        with sqlite3.connect(database) as db:
            log(
                f"Populating {database} using the contents of {len(workbook.sheetnames)} sheets found in {spreadsheet}."
            )

            # Iterate over the sheets in the workbook.
            for sheet_name in workbook.sheetnames:
                # Skip any sheets that were not explicitly requested.
                if sheet and sheet_name not in sheet:
                    log(f"Skipping sheet named '{sheet_name}'.")
                    continue

                # Reference the data in the rows of the current sheet.
                rows = workbook[sheet_name].iter_rows(values_only=True)

                # Create a table for each sheet.
                headings = list(next(rows))
                columns = get_column_names(sheet_name, headings, log)
                table_name = normalize(sheet_name)
                log(
                    f"Mapping contents of sheet '{sheet_name}' to table '{table_name}':"
                )

                # Determine whether any columns in this sheet were selected.
                selected = []
                index = 0
                for heading, column_name in zip(headings, columns):
                    if not column or heading in column or column_name in column:
                        selected.append(index)
                        log(f"  {heading} -> {column_name}")
                    index += 1

                # Only create the table if columns were selected.
                if selected:
                    columns = [columns[i] for i in selected]
                    create_table_sql = (
                        f"CREATE TABLE {table_name} ({', '.join(columns)})"
                    )
                    log(f"DB executing SQL: {create_table_sql}")
                    db.execute(create_table_sql)
                else:
                    log(
                        f"Skipping table {table_name} because no columns were selected."
                    )
                    continue

                # Insert the rows in batches.
                insert_rows = f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES ({', '.join(['?'] * len(selected))})"
                cursor = db.cursor()
                batch: list[tuple[str]] = []
                total = [0]

                def insert(row: tuple[str], batch_size: int = 1000) -> None:
                    """
                    Insert a row into the database.

                    Args:
                        row (tuple[str]): The row to be inserted.
                        batch_size (int, optional): The batch size for executing multiple rows at once. Defaults to 1000.

                    Returns:
                        None
                    """
                    flush = False if row else True

                    if row:
                        batch.append(row)

                    if flush or batch and len(batch) >= batch_size:
                        log(f"  ... inserting {len(batch)} rows")
                        cursor.executemany(insert_rows, batch)
                        total[0] += len(batch)
                        batch.clear()
                        if flush:
                            log(f"Writing {total[0]} rows...")
                            db.commit()
                            log("DONE!")

                # Insert rows in batches, flushing the final rows.
                for row in rows:
                    if row:
                        if column and selected:
                            insert([row[i] for i in selected])
                        else:
                            insert(row)
                else:
                    insert(None)  # Flush a partial batch.

    finally:
        # Clean up.
        workbook.close()
