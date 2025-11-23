# What is Lexerow ?
Lexerow is a backend dotnet library to process easily datarows and cells in Excel files (xlsx).

You can for example detect empty cell in a column and set a specific value. 
You can compare a cell value to a specific value and then put a new value in the same cell row or in another cell.

Lexerow is developed in C# and can be used in any dotnet application. 

Lexerow is an open source library.

# A quick example

## Problem: empty cells

You have an Excel file containing a datatable: the first line is the header, and others are datarows of the table.
In column B, some cells are empty, and it's a problem. It would better to have a value in each cell.


<p align="center">
    <img src="0-Docs/Readmemd/datarow_cells_empty_2025-04-13.jpg" alt="Some cells are empty">
</p>

So to put the value 0 in each empty cell in column B, Lexerow will help you to do that easily with few lines of code.

<p align="center">
    <img src="0-Docs/Readmemd/datarow_cells_set_zero_2025-04-13.jpg" alt="Cells have now values">
</p>

## The solution, in 2 stages

-1/ Create a dotnet program to execute your scripts.
 
-2/ Create a script to fix cell values in the Excel datatable. Create, Modify scripts to process data in your Excel file.


## The Script to fix values

To process datarow of the excel file as explained, Lexerow provide a powerful instruction which is: OnExcel ForEachRow If/Then.

Let's consider the excel file to fix blank values is "file.xlsx"
The first row is the header. Data starts at the second row which is the default case.

Create a basic script and save it "script.lxrw"

```
# process datarow of the Excel, one by one
OnExcel "file.xlsx"
    ForEachRow
	  If B.Cell=blank Then B.Cell=0
    Next
End OnExcel	
```

This script will scan each datarow present in the first sheet starting by defaut from the row #2 until the last one automatically.


This a very basic script, but of course it's possible to create more complex script to manage different cases.

## A C# program to execute the script

Create a program in C# and use the Lexerow library in this way:

```
// create the core engine
LexerowCore core = new LexerowCore();

// load and execute the script   
core.LoadExecScript("MyScript", "MyScript.lxrw");   
```

This is the minimum C# program you have to write.

# Package available on Nuget

Lexerow library is packaged as a nuget ready to use:

https://www.nuget.org/packages/Lexerow


# to go further with script

It's possible to set a different first data row. For example, start at the row #3 in place of default row #2.

```
# process datarow of the Excel, one by one
OnExcel "file.xlsx"
	FirstRow 3
    ForEachRow
	  If B.Cell=blank Then B.Cell=0
    Next
End OnExcel	
```

It's possible to check many cases in If instruction.

```
If A.Cell=12
If A.Cell>8.55
If A.Cell<>"Hello"
If A.Cell=blank
If A.Cell=null
```

In Then part, you can set a value to a cell.
Type of value can be: int, double, string.

To clear the cell value, you have to put blank. 
The formating of the cell will remain: background color and border.

To remove completly a cell, to have to set it to null.
 

```
Then A.Cell=13
Then A.Cell=25.89
Then A.Cell="Hello"
Then A.Cell=blank
Then A.Cell= null
```

Several instructions in Then part is also possible, example:

```
OnExcel "file.xlsx"
	FirstRow 3
    ForEachRow
	  If A.Cell="Y" Then 
		A.Cell="N"
		B.Cell=25.89
		C.Cell=blank
		End If
    Next
End OnExcel	
```

# Project Wiki

You can find more information on how use all available functions on the library here:

https://github.com/Pierlam/Lexerow/wiki


# Dependency

To access Excel content file, Lexerow uses the great NPOI library found on Nuget here:

https://www.nuget.org/packages/NPOI

NPOI source code is hosted on github here:

https://github.com/nissl-lab/npoi


