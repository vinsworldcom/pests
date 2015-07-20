NAME

PESTS - Parse Excel Spreadsheets To Single
Author:  Michael J. Vincent


DESCRIPTION

Script parses Excel (.XLS[X]) files in a given directory and extracts
the provided cells into a single output, one row for each parsed sheet
with the columns equal to the values of the cells in the provided
command line argument.


DEPENDENCIES

  Win32::OLE
  
This module uses the Win32::OLE interface to Excel to manipulate Excel 
files.
