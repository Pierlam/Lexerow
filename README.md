# What is Lexerow ?
Lexerow is a dotnet library to process very easily datarows and cells in Excel files.
You can for example detect empty cell in a column and set a specific value.

Example:
for each row, ff cell of column A is empty then set value 0.

Source Code:
```
   LexerowCore core = new LexerowCore();
   string fileName = "MyExcelFile.xlsx";
   ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);
todo:
...
```
