[![Build Status](https://travis-ci.org/martsberger/tablepyxl.svg?branch=master)](https://travis-ci.org/martsberger/tablepyxl)

# tablepyxl - A python library to convert html tables to Excel

## Introduction

Tablepyxl is a bridge between html tables and [openpyxl](http://openpyxl.readthedocs.org/en/default/).
If you can make an html table, you can make an Excel workbook.

## Usage

If your html table is in a string, you can write an Excel file with the `document_to_xl` function:
```
from tablepyxl import tablepyxl

table = "<table>"
        " <thead>"
        "  <tr>"
        "   <th>Header 1</th>"
        "   <th>Header 2</th>"
        "  </tr>"
        " </thead>"
        " <tbody>"
        "  <tr>"
        "   <td>Cell contents 1</td>"
        "   <td>Cell contents 2</td>"
        "  </tr>"
        " </tbody>"

tablepyxl.document_to_xl(table, "/path/to/file")
```

If your html table is in a file, read it into a string first:
```
from tablepyxl import tablepyxl

doc = open("/file/with/html/table", "r")
table = doc.read()

tablepyxl.document_to_xl(table, "/path/to/output")
```

Convert your html to an openpyxl workbook object instead of a file so that you can do further work:
```
from tablepyxl import tablepyxl

doc = open("/file/with/html/table", "r")
table = doc.read()

wb = tablepyxl.document_to_wb(table)

# For example, you can add another document to the same workbook
# in a new sheet:
doc2 = open("/file/with/html/table2", "r")
table2 = doc2.read()

wb = tablepyxl.document_to_wb(table, wb=wb)
```

Notes:
* A document with more than one table will write each table to a separate sheet
* Sheet names match the name attribute of the table element
* Multiple tables can be added to the same sheet using the `insert_table` method.

## Styling and Formatting

Tablepyxl intends to support all of the style and formatting options supported by Openpyxl. Here are the
currently supported styles:

### Font
* bold via the font-weight style, e.g. `<td style="font-weight:bold;">`
* color via the color style, e.g. `<td style="color:ff0000">`

### Alignment
* horizontal via the text-align style
* vertical via the vertical-align style
* wrap_text via the white-space style

### Fill
* Solid background color via the background-color style

### Border
* style and color for the top border via border-top-style and border-top-color styles

### Cell types
Cell types can be set by adding any of the following classes to the td element:
* TYPE_STRING
* TYPE_FORMULA
* TYPE_NUMERIC
* TYPE_BOOL
* TYPE_CURRENCY
* TYPE_NULL
* TYPE_INLINE
* TYPE_ERROR
* TYPE_FORMULA_CACHE_STRING
* TYPE_INTEGER

### Number formatting
* Currency is formatted using FORMAT_CURRENCY_USD_SIMPLE
* Dates are formatted using 'mm/dd/yyyy'
* Numeric values are formatted with commas every 3 digits if the commas are present in the html


### Merging
* Cells can be merged using the colspan and rowspan attributes of td elements

## License

MIT (http://opensource.org/licenses/MIT)

## Contributors

* [amehjabeen](https://github.com/amehjabeen)
* [bmdavi3](https://github.com/bmdavi3)
* [scottsexton](https://github.com/scottsexton)
