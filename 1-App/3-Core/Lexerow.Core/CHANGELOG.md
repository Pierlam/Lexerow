## 0.5.0 Release (2025-12-31)

-Access to Excel: NPOI Library replaced by OpenExcelSdk.

-New: Date function in place. Date(year, month, day)
an be used in If or then part, in a variable and more.
By default display with the default format: dd/mm/yyyy

-New: You can change the default display format for date
e.g. $DateFormat="yyyy-mm-dd

-Now string comparison are by default case insensitive.
But it's possible to change the default behavior by setting the system var to true.
$StrCompareCaseSensitive=true

-147 unit tests, all are green.

-Bugs fixed

