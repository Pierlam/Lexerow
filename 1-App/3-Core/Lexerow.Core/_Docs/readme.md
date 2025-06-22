# What is Lexerow ?

Lexerow is a backend dotnet library to process easily datarows and cells in Excel files.

For example you can detect empty cell in a column and set a specific value.
You can compare a cell value to a specific value and then put a new value in the same cell row or in another cell.

Lexerow is developed in C# and can be used in any dotnet application.

Lexerow is an open source library.

# A quick example

## Problem: empty cells

You have an Excel file containing a datatable in the first sheet: the first line is the header, and others are datarows of the table.
In column B, some cells are empty, and it's a problem to do calculation. It would better to have a value in each cell.

```
+------+-------+
|  Id  | Value |
+------+-------+
|   1  |   12  |
|   2  |       |  <= is empty!
|   3  |  234  |
|   4  |       |  <= is empty!
|   5  |  631  |
+------+-------+
```


So to put the value 0 in each empty cell in column B, Lexerow will help you to do that easily with some lines of code.

```
+------+-------+
|  Id  | Value |
+------+-------+
|   1  |   12  |
|   2  |    0  |  <= set to 0
|   3  |  234  |
|   4  |    0  |  <= set to 0
|   5  |  631  |
+------+-------+
```

## How it works

To proceed datarow as explained, Lexerow provide the main function which is: "OnExcel ForEachRow If-Then".

To understand how it works here is the pseudo code:

```
# open the excel file to process
file=OpenExcel("MyFile.xlsx")

# process datarow of the Excel, one by one
OnExcel file
  OnSheet 0,0
    ForEach Row
	  If B.Cell=null Then B.Cell= 0
    End
```

## How to implement using C# 

Create a program in C# and use the Lexerow library in this way:

```
LexerowCore core = new LexerowCore();
string fileName = "MyFile.xlsx";
   
// file= OpenExcel("MyExcelFile.xlsx")
core.Builder.CreateInstrOpenExcel("file", fileName);
   
// Comparison: B.Cell=null  (B -> index 1)
InstrCompColCellValIsNull instrCompIf = core.Builder.CreateInstrCompCellValIsNull(1);

// Set: B.Cell= 0
InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(1, 0);

// If B.Cell=null Then B.Cell= 0
InstrIfColThen instrIfColThen;
core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);

// OnExcel ForEach Row IfColThen, sheetNum=0, firstDataRow=1 below the header
core.Builder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);

// execute the instructions -> empty cells in col B will be remplaced by the value 0
core.Exec.Execute();
```


# Project Github 

The source code is hosted on github here:

https://github.com/Pierlam/Lexerow


# Project Wiki

It is possible to check many cell type in If instruction: IsNull, Int, Double, DateTime, DateOnly and also TimeOnly.

```
If A.Cell = null
If A.Cell = blank
If A.Cell = 12
If A.Cell = "tchao"
If A.Cell > 02/19/2025
If A.Cell < 01/02/2020 12:34:56
if A.Cell in ["yes", "y", "ok"]
```

You can put one or more Set Cell Value instruction in the Then part. 
Many type to set are available: Int, Double, DateTime, DateOnly and also TimeOnly.

It is also possible to remove the cell by setting null. 
Another option is to set Blank to a cell value, in this case the style of cell (BgColor, FgColor, Border,..) will remain.

```
Then A.Cell= 13
Then A.Cell= "Hello"
Then A.Cell= 12/04/2025
Then A.Cell= null
Then A.Cell= blank
```

You can find more information on how use all available functions on the library here:

https://github.com/Pierlam/Lexerow/wiki


# Dependency

To access Excel content file, Lexerow uses the great NPOI library found on Nuget here:

https://www.nuget.org/packages/NPOI

NPOI source code is hosted on github here:

https://github.com/nissl-lab/npoi


