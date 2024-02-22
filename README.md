# excel-utility-scripts-and-templates
A set of VBA, TypeScript and Python scripts useful for daily Excel automation projects, and Excel template files to demonstrate specific use cases of excel formulas and features.

## Repository structure
Each file type correspond to a specific type of script or template :
* [VBA scripts](https://learn.microsoft.com/en-us/office/vba/api/overview/) which are only compatible with desktop versions of Excel. These scripts are written in Visual Basic language and stored in ".bas" text files.
* [Office scripts](https://learn.microsoft.com/en-us/office/dev/scripts/develop/scripting-fundamentals) which are only compatible with web versions of Excel and some desktop versions (require consistent internet connection and that file is located on OneDrive). These scripts are written in TypeScript language and stored in ".ts" text files.
* Template or example excel files are compatible (with some [limitations](https://support.office.com/article/f0dc28ed-b85d-4e1d-be6d-5878005db3b6)) with both Excel desktop and [Excel for the web](https://learn.microsoft.com/en-us/office365/servicedescriptions/office-online-service-description/excel-online). These files have a ".xlsx" (for examples) or ".xltx" (for templates) extension and are a zip archive of XML files, so make sure to download file to preview it.
* [Python scripts](https://www.python.org/about/gettingstarted/) to easily automate files from outside. These scripts are stored in ".py" text files. Python is a free, easy and widespread programming language, with built in file management librairies.

## Installation
* VBA scripts : ".bas" files must be imported from the VBA IDE (from Excel desktop : alt + F11 > File > Import file).
* Office scripts : text contained in ".ts" files must be copy-pasted in the Office Scripts IDE. *Please note that, the Office scripts are stored here in ".ts" format and would require to be restructured to comply to ".osts" format (not only renamed).*
* Template or example Excel files : open the file with Excel.
* Python scripts : ".py" files must be run from a command line interface where the python programming language is installed.

## Features
The scripts available on this repository are listed and explained below.

### VBA - gather sheets summary
*related file : [GatherSheetsSummary.bas](https://github.com/ronan-deshays/excel-utility-scripts-and-templates/blob/main/GatherSheetsSummary.bas)*

For each sheet in a workbook, gather a range located always in the same cell on each sheet. This range contains e.g. a summary of data contained in the active sheet. So that the juxtaposition of ranges makes a summary of the whole workbook.

### Office - array form to database
*related file : [ArrayForm2Database.ts](https://github.com/ronan-deshays/excel-utility-scripts-and-templates/blob/main/ArrayForm2Database.ts)*

Build and update a database based on an array form.
More precisely, users fill the array form, and a script organize the data in a database (Excel table), which enables Pivot Tables or Power Platform usage of this data.

### Example - bypass pivot table data display limitation
*related file : [PivotTableDataDisplay.xlsx](https://github.com/ronan-deshays/excel-utility-scripts-and-templates/blob/main/PivotTableDataDisplay.xlsx)*

Excel pivot table feature obliges user to agregate data (using a sum or other functions), which is something impossible with text values or other specific types.

### Python - osts2ts - githubise Office Scripts
*related file : [osts2ts.py](https://github.com/ronan-deshays/excel-utility-scripts-and-templates/blob/main/osts2ts.py)*

Converts all non easy to read .osts files to readable .ts  files located in the same folder as the python script, and save them in a target folder of your choice. An OSTS file is created when saving an Office script to Onedrive, but it is saved as a JSON structure. This script helps converting this file to a "code editor friendly" file.

### Python - Excel Formula Parser - githubise Excel formulas
*related folder : [ExcelFormulaParser](https://github.com/ronan-deshays/excel-utility-scripts-and-templates/tree/master/ExcelFormulaParser)

Enables versioning of Excel formulas, writing a formula on multiple lines and adding comments.
The parsing process does the following :
* python-like comments management : comment is recognized as a line which begin with "# " (hastag followed by space)
* line breaks and spaces remove
* create one line per formula : formula is detected as a line break followed by "=" (equal) sign

**How to use :**
* add your formula to [excel_formula_in.txt](https://github.com/ronan-deshays/excel-utility-scripts-and-templates/blob/master/ExcelFormulaParser/excel_formula_in.txt)
* run [parser.py](https://github.com/ronan-deshays/excel-utility-scripts-and-templates/blob/master/ExcelFormulaParser/excel_formula_in.txt)
* get result in [excel_formula_out.txt](https://github.com/ronan-deshays/excel-utility-scripts-and-templates/blob/master/ExcelFormulaParser/excel_formula_in.txt)

**Known limitations :**
* variable name containing space not supported
* unable to make difference between formula begin and "=" sign in formula body
