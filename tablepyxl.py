from bs4 import BeautifulSoup as BS
from openpyxl import Workbook
from openpyxl.cell import Cell


def get_tables(doc):
    """
    Turn a string containing html into a list of tables
    """
    soup = BS(doc)
    return soup.find_all('table')


def write_rows(worksheet, elem, row, column=1):
    """
    Writes every tr child element of elem to a row in the worksheet
    elem could be a thead or tbody, so we write each th and td to a cell.

    returns the next row after all rows are written
    """
    for table_row in elem.find_all('tr'):
        for cell in table_row.find_all('th') + table_row.find_all('td'):
            worksheet.cell(row=row, column=column).value = cell.text
            column += 1
        row += 1
    return row


def table_to_sheet(table, wb):
    """
    Takes a table and workbook and writes the table to a new sheet.
    The sheet title will be the same as the table attribute name.
    """
    ws = wb.create_sheet(title=table.get('name', None))
    insert_table(table, ws, 1, 1)


def document_to_xl(doc, filename):
    """
    Takes a string representation of an html document and writes one sheet for
    every table in the document. The workbook is written out to a file called filename
    """
    wb = Workbook()
    wb.remove_sheet(wb.active)
    tables = get_tables(doc)

    for table in tables:
        table_to_sheet(table, wb)

    wb.save(filename)


def insert_table(table, worksheet, column, row):
    row = write_rows(worksheet, table.thead, row, column)
    row = write_rows(worksheet, table.tbody, row, column)


def insert_table_at_cell(table, cell):
    """
    Inserts a table at the location of an openpyxl Cell object.
    """
    ws = cell.parent
    column, row = cell.column, cell.row
    insert_table(table, ws, column, row)
