# Excelinfo

Gets index header names and column names from excel files.

This is used to gather information about the columns in a spreadsheet 
and use it to generate a mapping file, for information about how to process the spreadsheet.

Output is always a text file, with the following format:

HEADER = INDEX, ColumnName = COLUMN NAME
```
A = 0, ColumnName = Column A
B = 1, ColumnName = Column B
C = 2, ColumnName = Column C
```