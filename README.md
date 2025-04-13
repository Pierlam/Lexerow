# What is Lexerow ?
Lexerow is a backend dotnet library to process easily datarows and cells in Excel files (xlsx).

You can for example detect empty cell in a column and set a specific value. 
You can compare a cell value to a specific value and then put a new value in the same cell row or in another cell.

Lexerow is developed in C# and can be used in any dotnet application. 

Lexerow is an open source library.

# A quick example

## Problem: empty cells

You have an Excel file containing a datatable: the first line is the header, and others are datarows of the table.
In column B, some cells are empty, and it's a problem to do calculation. It would better to have a value in each cell.


<p align="center">
    <img src="0-Docs/Readmemd/datarow_cells_empty_2025-04-13.jpg" alt="Some cells are empty">
</p>

So to put the value 0 in each empty cell in column B, Lexerow will help you to do that easily with some lines of code.

<p align="center">
    <img src="0-Docs/Readmemd/datarow_cells_set_zero_2025-04-13.jpg" alt="Cells have now values">
</p>


## How to proceed

Create a program in C# and use the Lexerow library in this way:

```
LexerowCore core = new LexerowCore();
string fileName = "MyFile.xlsx";
   
// file= OpenExcel("MyExcelFile.xlsx")
core.Builder.CreateInstrOpenExcel("file", fileName);
   
//--Comparison: B.Cell=null  (B -> index 1)
InstrCompColCellValIsNull instrCompIf = core.Builder.CreateInstrCompCellValIsNull(1);

//--Set: B.Cell= 0
InstrSetCellVal instrSetValThen = core.Builder.CreateInstrSetCellVal(1, 0);

// If B.Cell=null Then B.Cell= 0
InstrIfColThen instrIfColThen;
core.Builder.CreateInstrIfColThen(instrCompIf, instrSetValThen, out instrIfColThen);

// OnExcel ForEach Row IfColThen, sheetNum=0, firstDataRow=1 below the header
core.Builder.CreateInstrOnExcelForEachRowIfThen("file", 0, 1, instrIfColThen);

// execute the instructions -> empty cells in col B will be remplaced by the value 0
core.Exec.Execute();
```

# Package available on Nuget

Lexerow library is packaged as a nuget ready to use:

https://www.nuget.org/packages/Lexerow#versions-body-tab

# Project Wiki

It is possible to check many cell type in If instruction: IsNull, Int, Double, DateTime, DateOnly and also TimeOnly.

You can put one or more Set Cell Value instruction in the Then part. Many type to set are available:
Null, Int, Double, DateTime, DateOnly and also TimeOnly.

You can find more information on how use all available functions on the library here:

https://github.com/Pierlam/Lexerow/wiki

# Dependency

To access Excel content file, Lexerow uses the great NPOI library found on Nuget here:

https://www.nuget.org/packages/NPOI

NPOI source code is hosted on github here:

https://github.com/nissl-lab/npoi


