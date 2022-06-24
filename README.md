# Directory Tree Utility

This was developed as part of my job to fit client requirements. Python program produces an Excel file which acts as a front-end to an external drive's directory tree.

## Overview

Documents are to be provided to client periodically via external drives. These documents are primarily taken from [Aconex](https://www.oracle.com/uk/industries/construction-engineering/aconex-project-controls/), am Oracle cloud file management and storage solution.

Client requested an Excel file to act as a front end to these drives.

Python program produces macro enabled Excel file, which includes a data tab listing entire directory tree structure, with hyperlinks to folders and files. Program also merges Aconex report to Excel file, providing metadata for each file.

Program is intended to be run on external hard drives, which would be provided to client, where drive letters could change. Script was developed with this in mind.



## Prerequisites

* Windows OS (will later be refactored to work with Mac\Linux. Issues are related to directory path formats)
* Python 3.9 installed
* Libraries installed from ***requirements.txt***
* Full document report from Aconex
	- This must include ***Document No***, ***File*** & ***File Name*** columns
	- Unnecessary header rows must be removed, so first row is the column headers
	- Rename file to ***metadata.xls*** and store in same folder as Python script
* Script will fail if number of files & folders in target directory exceeds limit


## Run Script

```bash
python 'handover_utility.py' '<target directory>'
```
	
## Files & Folders

### vbaProject folder & vbaProject.bin file

The newly produced Excel file must be in XLSM format (macro enabled). 

We must also add a macro to this file (so when we click on a folder in table of contents, it will search the data tab for that folder and navigate to it).

Unfortunately, the library we're using to write to Excel, [Xlswriter](https://xlsxwriter.readthedocs.io/working_with_macros.html), won't accept an XLSM extension.

We instead  write to `temp\temp.xlsx` and then change the name to `handover_utility.xlsm`.

However, Xlsxwriter won't save this as it doesn't contain a macro.

Instead, we create a macro in a separate workbook manually, and then extract the `VbaProject.bin` macro which contains the functions/macros we want.

```bash
python .vba_extract.py .vbaProject.xlsm
```

Assuming the outputted `vbaProject.bin` file is saved in the same directory as the `handover_utility.py` script, the macro will be inserted into our newly outputted Excel file before it's saved.

### handover_utilty.py

Main Python Script

### metadata.xls

Example of what metadata file could look like