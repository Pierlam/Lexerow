using Lexerow.Core.System;
using OpenExcelSdk.System;

namespace Lexerow.Core.Utils;

public class ExcelExtendedUtils : ExcelUtils
{
    /// <summary>
    /// Does the cell type match the If-Comparison cell.Value type?
    /// used for: If A.Cell=10
    /// </summary>
    /// <param name="cellType"></param>
    /// <param name="instr"></param>
    /// <returns></returns>
    public static bool MatchCellTypeAndIfComparison(ExcelCellType cellType, ValueBase value)
    {
        if (cellType == ExcelCellType.String && value.ValueType == System.ValueType.String)
            return true;

        if (cellType == ExcelCellType.Double && value.ValueType == System.ValueType.Double)
            return true;

        if (cellType == ExcelCellType.Integer && value.ValueType == System.ValueType.Int)
            return true;

        // int - double
        if (cellType == ExcelCellType.Integer && value.ValueType == System.ValueType.Double)
            return true;

        // double - int
        if (cellType == ExcelCellType.Double && value.ValueType == System.ValueType.Int)
            return true;

        if (cellType == ExcelCellType.DateOnly && value.ValueType == System.ValueType.DateOnly)
            return true;

        if (cellType == ExcelCellType.DateTime && value.ValueType == System.ValueType.DateTime)
            return true;

        if (cellType == ExcelCellType.TimeOnly && value.ValueType == System.ValueType.TimeOnly)
            return true;

        // specific cases
        if (cellType == ExcelCellType.DateTime && value.ValueType == System.ValueType.DateOnly)
            return true;
        if (cellType == ExcelCellType.DateOnly && value.ValueType == System.ValueType.DateTime)
            return true;

        return false;
    }

    /// <summary>
    /// Does the cell type match the If-Comparison cell.Value type?
    /// used for: If A.Cell=B.Cell
    /// </summary>
    /// <param name="cellType"></param>
    /// <param name="instr"></param>
    /// <returns></returns>
    public static bool MatchCellTypeAndIfComparison(ExcelCellType cellType, ExcelCellType cellTypeB)
    {
        if (cellType == ExcelCellType.String && cellTypeB == ExcelCellType.String)
            return true;

        if (cellType == ExcelCellType.Double && cellTypeB == ExcelCellType.Double)
            return true;
        if (cellType == ExcelCellType.Integer && cellTypeB == ExcelCellType.Integer)
            return true;

        // int and double cases
        if (cellType == ExcelCellType.Integer && cellTypeB == ExcelCellType.Double)
            return true;
        if (cellType == ExcelCellType.Double && cellTypeB == ExcelCellType.Integer)
            return true;

        if (cellType == ExcelCellType.DateOnly && cellTypeB == ExcelCellType.DateOnly)
            return true;

        if (cellType == ExcelCellType.DateTime && cellTypeB == ExcelCellType.DateTime)
            return true;

        if (cellType == ExcelCellType.TimeOnly && cellTypeB == ExcelCellType.TimeOnly)
            return true;

        // specific cases
        if (cellType == ExcelCellType.DateTime && cellTypeB == ExcelCellType.DateOnly)
            return true;
        if (cellType == ExcelCellType.DateOnly && cellTypeB == ExcelCellType.DateTime)
            return true;

        return false;
    }
}