xlrd
====
xlrd is a library for reading data and formatting information from Excel
files in the historical ``.xls`` format.

.. warning::

  This library will no longer read anything other than ``.xls`` files. For
  alternatives that read newer file formats

The following are also not supported but will safely and reliably be ignored:

*   Charts, Macros, Pictures, any other embedded object, **including** embedded worksheets.
*   VBA modules
*   Formulas, but results of formula calculations are extracted.
*   Comments
*   Hyperlinks
*   Autofilters, advanced filters, pivot tables, conditional formatting, data validation

Password-protected files are not supported and cannot be read by this library.

Quick start:

.. code-block:: python

    import xlrd
    book = xlrd.open_workbook("myfile.xls")
    print("The number of worksheets is {0}".format(book.nsheets))
    print("Worksheet name(s): {0}".format(book.sheet_names()))
    sh = book.sheet_by_index(0)
    print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
    print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
    for rx in range(sh.nrows):
        print(sh.row(rx))

From the command line, this will show the first, second and last rows of each sheet in each file:

.. code-block:: bash

    python PYDIR/scripts/runxlrd.py 3rows *file_title*.xls
    
Details
==
*   Handling of Unicode
*   Dates in Excel spreadsheets
*   Named references, constants, formulas, and macros
*   Formatting information in Excel Spreadsheets
*   Loading worksheets on demand
*   API Reference
