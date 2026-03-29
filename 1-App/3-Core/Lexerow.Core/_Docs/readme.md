# What is Lexerow ?

Lexerow is a backend dotnet library providing a micro-language to process easily datarows and cells in Excel files.

A very basic use case should be to detect empty cell in a column and set a specific value.
You can compare a cell value to a specific value and then put a new value in the same row or in another cell of the row.

Lexerow is developed in C# and can be used in any dotnet application.

Lexerow is a free and open source library.

# A quick example

## Problem: empty cells


You have an Excel file containing a datatable in the first sheet: the first line is the header, and others are datarows of the table.
In column B, some cells are empty, and it's a problem. It would better to have a value in each cell.

Comment: In Excel language, we say blank rather than empty.

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


So to put the value 0 in each empty cell in column B, Lexerow will help you to do that easily with few lines of code.

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

## The solution

-1/ Create a script to fix cell values in the Excel datatable.

-2/ Create a dotnet program to execute your script.

-3/ Feel free to modify the script as needed and execute it again.
 

## The Script to fix values

To process datarow of the excel file as explained, Lexerow provide a powerful instruction which is: OnExcel.

Let's consider the excel file to fix blank values is "file.xlsx".
The first row is the header. Data starts at the second row which is the default case.

Create a basic script as a text file and save it let's say with this name: "script.lxrw"

The file name extension is free.

```
# process datarow of the Excel, one by one
OnExcel "file.xlsx"
    ForEachRow
	  If B.Cell=blank Then B.Cell=0
    Next
End OnExcel	
```

The script will scan each datarow present in the sheet starting by defaut from the row #2.
Each time the cell value in column B is blank, the value 0 will be set in place.
The execution will stop automatically after the last row was processed.

This a very basic script with few instructions to manage this standard case, but of course it's possible to create more complex scripts to manage all your specific cases.


## A C# program to execute the script

Now let's open your Visual Studio or VSCode.
First import the nuget Lexerow package in your solution or project. 

Create a program in C# and use it in this way:


```
// create the core engine
LexerowCore core = new LexerowCore();

// load and execute the script   
core.LoadExecScript("MyScript", "script.lxrw");   
```

This is the smallest C# program you have to write.


# Project Github 

The source code is hosted on github here:

https://github.com/Pierlam/Lexerow


# To go further - Script tuning

## basic script sample

Now let's manage a specific case of your datatable.

If the header use 2 rows or more, it's possible to set a different first data row to process. For example, start at the row #3 in place of default row #2. You may use the instruction FirstRow.

```
# process datarow of the Excel, one by one
OnExcel "file.xlsx"
	FirstRow 3
    ForEachRow
	  If B.Cell=blank Then B.Cell=0
    Next
End OnExcel	
```

## If condition in script

There are several kind of checks available in If condition:

```
If A.Cell=12
If A.Cell>8.55
If A.Cell<>"Hello"
If A.Cell=blank
If A.Cell=null
If A.Cell <= Date(2019,11,14) 
If A.Cell >10 And B.Cell< 20
If (A.Cell >10 And B.Cell< 20) Or C.Cell=12
```

Remark: string comparison are case insensitive by default.
It's possible to modify this behaviour by putting this instruction at the top of the script:

```
$StrCompareCaseSensitive=true
```


## Then instruction in script

In Then instruction, you can set a value to a cell or clear it.

To clear the cell value, you can put blank in it. 
The formating of the cell will remain: background/foreground color and border.

To remove completly a cell (value and formatting), you have to set it to null. 


```
Then A.Cell=13
Then A.Cell=25.89
Then A.Cell="Hello"
Then A.Cell=blank             # cell formatting will stay
Then A.Cell=null              # cell formatting will be cleared
Then A.Cell=Date(2019,11,14)  # put a date value in the cell with the default format dd/mm/yyyy
Then CopyRow(fileRes)         # copy the current row to another excel file
```

By default date are displayed with the default format: dd/mm/yyyy
You can change the default display format for date by putting this instruction at the top of the script:

```
# would like to change the output format of dates
$DateFormat="yyyy-mm-dd"
```

For If and Then instruction, type of value can be: int, double, string, and Date.

DateTime, Time and currency are not yet managed.

You may want to apply many Then instructions for one If instruction; for example:


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

## More possible cases!

There are a lot of cases which are possible to put in script.

Take a look to samples:  https://github.com/Pierlam/Lexerow/wiki/Script-samples

# Manage errors

From a technical point of view, you can manage easily errors occured during compilation or during the execution of scripts.


```

// load and execute the script   
var result= core.LoadExecScript("MyScript", "MyScript.lxrw");   
if(!result.Res)
{
	// errors occured -> see result.ListError
	// On an error, see ErrorCode to know the type of error that occured
}
```

# Project Wiki

You can find more information on how to create more powerful scripts:

https://github.com/Pierlam/Lexerow/wiki


# Dependency

To access Excel content file, Lexerow uses now the OpenExcelSdk library available on Nuget here:

https://www.nuget.org/packages/OpenExcelSdk 

Source code is hosted on github here:

https://github.com/Pierlam/OpenExcelSdk



