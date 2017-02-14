PRE-REQUISITES
- Requires Python 3.6.0 ( Download it from https://www.python.org/downloads/release/python-360 )
- Requires openpyxl library for reading Excel file ( Find it in the folder )

HOW TO INSTALL THE LIBRARY
- After you install Python 3.6.0, a folder by the name Python36 gets created.
- Find the library in the folder, copy in the Python36 folder and run command "pip install openpyxl" in Command Prompt to install

ENVIRONMENT SETUP
- Install Python 3.6.0 and openpyxl library before running the script
- It is required that error-code-ref-complete.xlsx be stored in the same folder as the script file

HOW TO RUN THE SCRIPT FILE
- Once the library is installed, run the file using the command "python file10.py" in Command Prompt

WHAT THIS SCRIPT DOES
- The scipt file is named file10.py
- This script reads sheets error-code-ref-general and error-code-ref-upgrade of Excel workbook error-code-ref-complete.xlsx
- This script creates two HTML files by name "error-code-ref-general.html" and "error-code-ref-upgrade.html". Find these files in the same folder as the script and Excel file.
- It is required that the Excel file is of .xlsx format. Python library openpyxl does not work for .xls format.