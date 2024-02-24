"""xlsql converts an Excel .xlsx spreadsheet into a sqlite3 database.

Usage: xlsql [OPTIONS] SPREADSHEET

  Convert an Excel spreadsheet into a SQLite database.

  Args:     spreadsheet (str): The path to the Excel spreadsheet.     database
  (str): The name of the database to create.     force (bool): Flag to
  overwrite an existing database.

Options:
  --database PATH  The name of the database to create.
  --force          Overwrite an existing database.
  --help           Show this message and exit.

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


@click.command()
@click.argument(
    "spreadsheet", type=click.Path(exists=True, readable=True, dir_okay=False)
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
def main(spreadsheet, database, force) -> None:
    """
    Convert an Excel spreadsheet into a SQLite database.

    Args:
        spreadsheet (str): The path to the Excel spreadsheet.
        database (str): The name of the database to create.
        force (bool): Flag to overwrite an existing database.

    Returns:
        None
    """
    # Ensure that the target database won't be overwritten, or that it's OK to
    # overwrite it.
    existing = Path(database)
    if database and existing.exists() and existing.stat().st_size:
        print(f"{database} exists!")
        if force:
            print("Overwriting", database)
            existing.unlink()
        else:
            raise click.ClickException(
                f"Cowardly refusing to overwrite existing db: {database}"
            )

    try:
        # Load the spreadsheet.
        workbook = openpyxl.load_workbook(spreadsheet)

        # Create a new SQLite database and connect to it.
        with sqlite3.connect(database) as db:
            # Iterate over the sheets in the workbook.
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                rows = sheet.iter_rows(values_only=True)

                # Create the table.
                headings = [normalize(heading) for heading in next(rows)]
                table_name = normalize(sheet_name)
                print(table_name, headings)
                db.execute(f"CREATE TABLE {table_name} ({', '.join(headings)})")

                # Insert the rows in batches.
                insert_rows = f"INSERT INTO {table_name} ({', '.join(headings)}) VALUES ({', '.join(['?'] * len(headings))})"
                cursor = db.cursor()
                batch: list[tuple[str]] = []

                def insert(row: tuple[str], batch_size: int = 1000) -> None:
                    """
                    Insert a row into the database.

                    Args:
                        row (tuple[str]): The row to be inserted.
                        batch_size (int, optional): The batch size for executing multiple rows at once. Defaults to 1000.

                    Returns:
                        None
                    """
                    print(row)
                    flush = False if row else True

                    if row:
                        batch.append(row)

                    if flush or batch and len(batch) >= batch_size:
                        cursor.executemany(insert_rows, batch)
                        batch.clear()
                        if flush:
                            db.commit()

                # Insert rows in batches, flushing the final rows.
                for row in rows:
                    if row:
                        insert(row)
                else:
                    insert(None)

    finally:
        # Clean up.
        workbook.close()


if __name__ == "__main__":
    main()
