## 0.6.0 Release (2026-02-10)

todo

## 0.5.0 Release (2026-01-10)

-Access to Excel: NPOI Library replaced by OpenExcelSdk.

-New: Date function in place. Date(year, month, day)

an be used in If or then part, in a variable and more.

By default display with the default format: dd/mm/yyyy

-New: You can change the default display format for date:

e.g. $DateFormat="yyyy-mm-dd"

-Now string comparison are by default case insensitive.

But it's possible to change the default behavior by setting the system var to true.

$StrCompareCaseSensitive=true

-And and Or boolean operator can be used in If conditiion.

e.g.  If A.Cell>10 And A.Cell<20 Then..

-Variable can be used. e.g. 

a=12

..

If A.Cell>a Then

-Creation of the class Diagnotics available in the core to manage easily logs/traces.

-166 unit tests, all are green.

-Bugs fixed

