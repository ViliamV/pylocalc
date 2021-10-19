# PyLOcalc

[![forthebadge](https://forthebadge.com/images/badges/0-percent-optimized.svg)](https://forthebadge.com)
[![forthebadge](https://forthebadge.com/images/badges/made-with-python.svg)](https://forthebadge.com)
[![forthebadge](https://forthebadge.com/images/badges/you-didnt-ask-for-this.svg)](https://forthebadge.com)

Python interface for manipulating LibreOffice Calc spreadsheets

**DISCLAIMER: This is not production software! Backup your document before trying it!**

## About

LibreOffice/OpenOffice has API for many languages including Python, thanks to the Universal Network Objects (UNO).

_But its API is all but [Pythonic](https://docs.python.org/3/glossary.html)!_

I took inspiration from [this article](https://christopher5106.github.io/office/2015/12/06/openoffice-libreoffice-automate-your-office-tasks-with-python-macros.html)
and created simple wrapper around this API.

PyLOcalc also automatically opens a headless LibreOffice Calc document with basic read, write, and save functionality.
Therefore, it can be used as a library for other scripts that manipulate spreadsheets.

## Requirements

1. **python>=3.8**
2. You have to have LibreOffice with "python support" installed.
   In Ubuntu it's (apparently) `sudo apt install libreoffice-script-provider-python`,
   in Arch Linux I didn't have to do anything.

   You can try running `python -m uno` and it should not produce any output or error.
3. Macro security - Medium. Follow [this article](https://christopher5106.github.io/office/2015/12/06/openoffice-libreoffice-automate-your-office-tasks-with-python-macros.html).

## Installation

```bash
pip install pylocalc
```

## Basic usage

```python
import pylocalc

doc = pylocalc.Document('path/to/calc/spreadsheet.ods')
# You have to connect first
doc.connect()

# Get the sheet by index
sheet = doc[2]
# Or by name
sheet = doc[doc.sheet_names[1]]

# Get the cell by index
cell = sheet[10, 14]
# Or by "name"
cell = sheet['B12']

# Read and set cell value
print(cell.value)
> 'Some value'

cell.value = 12.2
print(cell.value)
> '12.2'

cell.value = 'Other value'
print(cell.value)
> 'Other value'

# Don't forget to save and close the document!
doc.save()
doc.close()
```

## Append rows and columns

PyLOcalc can append row and column values to the first available row or column.
It looks at the cell at the `offset` (default 0) and if the cell is empty it adds values there.

```python
import decimal
import pylocalc

doc = pylocalc.Document('path/to/calc/spreadsheet.ods')
doc.connect()
sheet = doc['Totals']

sheet.append_row(('2021-01-01', 123, 12.3, decimal.Decimal("0.111"), 'Yaaay'), offset=1)

sheet.append_column(('New column header'))

doc.save()
doc.close()
```

## Context manager

PyLOcalc `Document` can be used as context manager that automatically connects and closes the document.
If no error is raised in the context block it also **saves the document**.

```python
import pylocalc
with pylocalc.Document('path/to/calc/spreadsheet.ods') as doc:
    doc[0][1,10].value = 'I ❤️ context managers'
```
