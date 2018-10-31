## exporting excel with human-readable column widths on file open by human, let's do this!
When users open Excel files exported by python, the column widths are often left at defaults which means much of the text is hidden and the file is not immediately readable in Excel without the user adjusting the columns.  Granted, this is as simple a `select all the columns and double click on column widths`, so Excel and auto-adjust the columns... but should not we just take care of this in advance with software?  We think so, thus this library.

###How it works:
1- ingest an arbitrary excel file,
2- check all the cells in a single column, find the maximum length of the content
3- set the column width equal to this maximum length
4- repeat for all the columns

###Some references:
https://pypi.org/project/XlsxWriter/
