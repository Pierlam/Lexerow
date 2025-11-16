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


# Project Github 

The source code is hosted on github here:

https://github.com/Pierlam/Lexerow


# Project Wiki

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

You can find more information on how use all available functions on the library here:

https://github.com/Pierlam/Lexerow/wiki


# Dependency

To access Excel content file, Lexerow uses the great NPOI library found on Nuget here:

https://www.nuget.org/packages/NPOI

NPOI source code is hosted on github here:

https://github.com/nissl-lab/npoi


