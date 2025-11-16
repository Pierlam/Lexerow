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

## The solution, in 2 stages

-Create a script to fix cell values in the Excel datatable.

-Execute the script in a C# program.
 

## The Script to fix values

To process datarow of the excel file as explained, Lexerow provide a powerful instruction which is: OnExcel ForEachRow If/Then.

Let's consider the excel file to fix blank values is "MyFile.xlsx"

Create a basic script and save it "MyScript.lxrw"

```
# process datarow of the Excel, one by one
OnExcel "MyFile.xlsx"
    ForEachRow
	  If B.Cell=null Then B.Cell=0
    Next
End OnExcel	
```

This a very basic script, but of course it's possible to create more complex script to manage your cases.

## A C# program to execute the script

Create a program in C# and use the Lexerow library in this way:

```
// create the core engine
LexerowCore core = new LexerowCore();

// load and execute the script   
core.LoadExecScript("MyScript", MyScript.lxrw);   
```

This is the minimum C# program you have to write.

# Package available on Nuget

Lexerow library is packaged as a nuget ready to use:

https://www.nuget.org/packages/Lexerow#versions-body-tab

# Project Wiki

It is possible to check many cases in If instruction.

```
If A.Cell=12
If A.Cell>8.55
If A.Cell<>"Hello"
If A.Cell=blank
If A.Cell=null
```

You can put one or more Set Cell Value instruction in the Then part. 

It is also possible to remove the cell by setting null. 
Another option is to set Blank to a cell value, in this case the style of cell (BgColor, FgColor, Border,..) will remain.

```
Then A.Cell=13
Then A.Cell=25.89
Then A.Cell="Hello"
Then A.Cell=blank
Then A.Cell= null
```

You can find more information on how use all available functions on the library here:

https://github.com/Pierlam/Lexerow/wiki

# Dependency

To access Excel content file, Lexerow uses the great NPOI library found on Nuget here:

https://www.nuget.org/packages/NPOI

NPOI source code is hosted on github here:

https://github.com/nissl-lab/npoi


