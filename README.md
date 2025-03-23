# What is Lexerow ?
Lexerow is a backend dotnet library to process very easily datarows and cells in Excel files.
You can for example detect empty cell in a column and set a specific value.

Lexerow is developed in C# and can be used in any dotnet application.

Example:

For each row, if cell of column A is empty then set value 0.

Source Code:
```
   LexerowCore core = new LexerowCore();
   string fileName = "MyExcelFile.xlsx";
   ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);
todo:
...
```
