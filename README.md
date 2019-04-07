## exporting excel with human-readable column widths on file open by human, let's do this!
When users open Excel files exported by python, the column widths are often left at defaults.
 
 This means much of the text is hidden and the file contents are not easily readable in Excel without the user manually adjusting the columns.  
 
 Granted, this is as simple a `select all the columns and double click on column widths`, so Excel and auto-adjust the columns... but should not we just take care of this in advance with software?  We think so, thus this library.
 

###How it works:
1- ingest an arbitrary excel file,
2- check all the cells in a single column, find the maximum length of the content
3- set the column width equal to this maximum length
4- repeat for all the columns

### Usage:
Case 1 - show changes to sample excel file provided in repo
1- Open `xlsx_autofit_columns` excel file in Excel, visually note column widths.
2- Close excel file
3- From your terminal, run `xlsx_autofit_columns` using default values
Two options for acceptable python syntax:
python -m xlsx_autofit_columns
python xlsx_autofit_columns.py

4- Open excel file, note changed column widths
5- Alternatively, just look at changed timestamp in the directory


Case 2 - pass new filename as parameter to the script
example:
1- From your terminal, run:
python -m xlsx_autofit_columns sample_excel_data_WITH_FORMULA.xlsx
2- Confirm file changes by either inspecting timestamp changes, or opening file as in example above


### Sample Results

Sample Excel file from repo before running `xlsx_autofit_columns`
![Excel BEFORE column width fix](https://github.com/surfaceowl/xlsx_autofit_columns/blob/master/readme_images/excel.sample_before.png) 

...and after:
![Excel AFTER column width fix](hhttps://github.com/surfaceowl/xlsx_autofit_columns/blob/master/readme_images/excel.sample_after.png)

###References:
https://openpyxl.readthedocs.io/en/stable/index.html
