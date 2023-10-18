title: Working with Excel Files in Python
save_as: index.html
url:

This site contains pointers to the best information available about working with [Excel](https://products.office.com/en-us/excel) files in the [Python](http://www.python.org/) programming language.


## Reading and Writing Excel Files

There are python packages available to work with Excel files that will run on any Python platform and that do not require either Windows or Excel to be used. They are fast, reliable and open source:

### openpyxl

The recommended package for reading and writing Excel 2010 files (ie: .xlsx)

[Download](http://pypi.python.org/pypi/openpyxl) | [Documentation](https://openpyxl.readthedocs.org/) | [Bitbucket](https://bitbucket.org/openpyxl/openpyxl)

### pyexcel

Library with a unified api for reading and writing files with older Excel format (xls), Excel 2010 format (xlsx), OpenDocument Spreadsheet format (ods), and many others. 

[Download](https://pypi.org/project/pyexcel/) | [Documentation](https://docs.pyexcel.org/en/latest/) | [Bitbucket](https://github.com/pyexcel/pyexcel)

### xlsxwriter

An alternative package for writing data, formatting information and, in particular, charts in the Excel 2010 format (ie: .xlsx)

[Download](https://pypi.python.org/pypi/XlsxWriter) | [Documentation](https://xlsxwriter.readthedocs.org/) | [GitHub](https://github.com/jmcnamara/XlsxWriter)

### pyxlsb

This package allows you to read Excel files in the `xlsb` format.

[Download](https://pypi.org/project/pyxlsb) | [GitHub](https://github.com/willtrnr/pyxlsb)

### pylightxl

This package allows you to read `xlsx` and `xlsm` files and write `xlsx` files.

[Download](https://pypi.org/project/pylightxl) | [Documentation](https://pylightxl.readthedocs.io/en/latest/) | [GitHub](https://github.com/PydPiper/pylightxl)

### xlrd

This package is for reading data and formatting information from older Excel files (ie: .xls)

[Download](http://pypi.python.org/pypi/xlrd) | [Documentation](http://xlrd.readthedocs.io/en/latest/) | [GitHub](https://github.com/python-excel/xlrd)

### xlwt

This package is for writing data and formatting information to older Excel files (ie: .xls)

[Download](http://pypi.python.org/pypi/xlwt) | [Documentation](http://xlwt.readthedocs.io/en/latest/) | [Examples](https://github.com/python-excel/xlwt/tree/master/examples) | [GitHub](https://github.com/python-excel/xlwt)

### xlutils

This package collects utilities that require both `xlrd` and `xlwt`, including the ability to copy and modify or filter existing excel files.

***NB:** In general, these use cases are now covered by openpyxl!*

[Download](http://pypi.python.org/pypi/xlutils) | [Documentation](http://xlutils.readthedocs.io/en/latest/) | [GitHub](https://github.com/python-excel/xlutils)

## Writing Excel Add-Ins

The following products can be used to write Excel add-ins in Python. Unlike the reader and writer packages, they require an installation of Microsoft Excel.

### PyXLL

PyXLL is a commercial product that enables writing Excel add-ins in Python with no VBA. Python functions can be exposed as
worksheet functions (UDFs), macros, menus and ribbon tool bars.

[Homepage](https://www.pyxll.com) | [Features](https://www.pyxll.com/features.html) | [Documentation](https://www.pyxll.com/docs/index.html) | [Download](https://www.pyxll.com/download.html)

### xlwings

xlwings is an open-source library to automate Excel with Python instead of VBA and works on Windows and macOS: you can call Python from Excel and vice versa and write UDFs in Python (Windows only). xlwings PRO is a commercial add-on with additional functionality.

[Homepage](https://www.xlwings.org) | [Documentation](https://docs.xlwings.org/en/stable/) | [GitHub](https://github.com/xlwings/xlwings) | [Download](https://pypi.org/project/xlwings/)


## The Mailing List / Discussion Group

There is a [Google Group](http://groups.google.com/group/python-excel) dedicated to working with Excel files in Python, including the libraries listed above along with manipulating the Excel application via COM. 

## Commercial Development

The following companies can provide commercial software development and consultancy and are specialists in working with Excel files in Python:

<div class="affiliate-links"></div>

| [![Clark Consulting & Research]({static}/images/ccr_python_excel.png)](http://www.clark-consulting.eu/) |
| ------------------------------------------------------------------------------------------------------- |
| [Clark Consulting & Research](http://www.clark-consulting.eu/)                                          |
