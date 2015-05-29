from bs4 import BeautifulSoup as BS
from openpyxl import Workbook


# Turn a string containing html into a list of tables
def get_tables(doc):
    soup = BS(doc)
    return soup.find_all('table')


def write_rows(ws, elem, row):
    for table_row in elem.find_all('tr'):
        column = 1
        for cell in table_row.find_all('th') + table_row.find_all('td'):
            ws.cell(row=row, column=column).value = cell.text
            column += 1
        row += 1
    return row


def table_to_sheet(table, wb):
    ws = wb.create_sheet(title=table.get('name', None))
    row = write_rows(ws, table.thead, 1)
    row = write_rows(ws, table.tbody, row)


def document_to_xl(doc, filename):
    wb = Workbook()
    wb.remove_sheet(wb.active)
    tables = get_tables(doc)

    for table in tables:
        table_to_sheet(table, wb)

    wb.save(filename)