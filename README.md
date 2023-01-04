# excel-utility-scripts
A set of common VBA and TypeScript scripts useful for daily Excel automation projects.

## Repository structure
There are two types of scripts on this repository :

* [VBA scripts](https://learn.microsoft.com/en-us/office/vba/api/overview/) which are only compatible with desktop versions of Excel. These scripts are written in Visual Basic language and stored in ".bas" text files.
* [Office scripts](https://learn.microsoft.com/en-us/office/dev/scripts/develop/scripting-fundamentals) which are only compatible with web versions of Excel and some desktop versions (require consistent internet connection and that file is located on OneDrive). These scripts are written in TypeScript language and stored in ".ts" text files.

## Installation
* VBA scripts must be imported from the VBA IDE (from Excel desktop : alt + F11)
* Office scripts on this GitHub repository can be copy-pasted in the Office Scripts IDE. Alternatively, you can deposit them on folder usally located in : your-OneDrive/Documents/Office scripts. Please note that, for development purposes, the Office scripts are stored in this repository in ".ts" format and would require to be restructured to comply to ".osts" format (not only renamed), so the copy-paste method explained above is recommended.

## Features
The scripts available on this repository are listed and explained below.

### VBA - gather sheets summary
*related file : GatherSheetsSummary.bas*
For each sheet in a workbook, gather a range located always in the same cell on each sheet. This range contains e.g. a summary of data contained in the active sheet. So that the juxtaposition of ranges makes a summary of the whole workbook.

### Office - array form to database
*related file : ArrayForm2Database.ts*
Build and update a database based on an array form.
More precisely, users fill the array form, and a script organize the data in a database (Excel table), which enables Pivot Tables or Power Platform usage of this data.
