
import click
import openpyxl

@click.command()
@click.argument('spreadsheet', type=click.Path(exists=True))
def main(spreadsheet):

    # Load the spreadsheet
    workbook = openpyxl.load_workbook(spreadsheet)

    # Iterate over the sheets
    for sheet_name in workbook.sheetnames:
        print(sheet_name)
        sheet = workbook[sheet_name]
        # Do something with the sheet
        for row in sheet.iter_rows(values_only=True):
            print(row)
        print()

    # Close the workbook
    workbook.close()


if __name__ == '__main__':
    main()
