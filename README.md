# Lexerow
dotnet library to process easily datatables rows and cells in Excel files.

Example, pseudo-code:

```
ForEach Row
  Excel="Greater50Set12.xlsx"  Sheet=0 FirstRow=0
  If D.Cell>50 Then D.Cell=12
End
```

Source Code:
```
   LexerowCore core = new LexerowCore();
   string fileName = "Greater50Set12.xlsx";
   ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);
todo:
...
```
