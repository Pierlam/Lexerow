namespace Lexerow.Core.System;

/// <summary>
/// List of Excel columns or range of columns.
/// </summary>
public class ExcelListColsOrRangeCols
{
    /// <summary>
    /// Excel column or range of columns (start-end).
    /// </summary>
    public List<ExcelColOrRangeColsBase> ListRangeCols { get; set; } = new List<ExcelColOrRangeColsBase>();

    public bool ContainsCol(int colIndex)
    {
        if (colIndex < 1) return false;

        foreach (var colOrRangeCols in ListRangeCols)
        {
            // is it a col?
            ExcelCol excelCol = colOrRangeCols as ExcelCol;
            if (excelCol != null)
            {
                if (colIndex == excelCol.ColumnInt)
                    return true;

                if (colIndex < excelCol.ColumnInt)
                    // the defined col is greater
                    return false;
            }

            // is it a range of cols?
            ExcelRangeCols excelRangeCols = colOrRangeCols as ExcelRangeCols;
            if (excelRangeCols != null)
            {
                if (colIndex >= excelRangeCols.ColumnStartInt && colIndex <= excelRangeCols.ColumnEndInt)
                    return true;

                if (colIndex < excelRangeCols.ColumnStartInt)
                    // the defined col is greater
                    return false;
            }
        }

        return false;
    }
}

/// <summary>
/// Excel column or range of columns, base class.
/// </summary>
public abstract class ExcelColOrRangeColsBase
{
}

/// <summary>
/// A range of Excel columns, exp: A-C
/// </summary>
public class ExcelRangeCols : ExcelColOrRangeColsBase
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="columnStart"></param>
    /// <param name="columnStartInt"></param>
    /// <param name="columnEnd"></param>
    /// <param name="columnEndInt"></param>
    public ExcelRangeCols(string columnStart, int columnStartInt, string columnEnd, int columnEndInt)
    {
        ColumnStart = columnStart;
        ColumnStartInt = columnStartInt;
        ColumnEnd = columnEnd;
        ColumnEndInt = columnEndInt;
    }

    public string ColumnStart { get; set; }
    public int ColumnStartInt { get; set; }

    public string ColumnEnd { get; set; }
    public int ColumnEndInt { get; set; }
}

/// <summary>
/// An excel column, exp A.
/// </summary>
public class ExcelCol : ExcelColOrRangeColsBase
{
    public ExcelCol(string column, int columnInt)
    {
        Column = column;
        ColumnInt = columnInt;
    }

    /// <summary>
    /// excel Column letter, exp: A
    /// </summary>
    public string Column { get; set; }

    /// <summary>
    /// excel column number, for A=1
    /// </summary>
    public int ColumnInt { get; set; }
}