using Lexerow.Core.System;

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
    public static bool MatchCellTypeAndIfComparison(CellRawValueType cellType, ValueBase value)
    {
        if (cellType == CellRawValueType.String && value.ValueType == System.ValueType.String)
            return true;

        if (cellType == CellRawValueType.Numeric && value.ValueType == System.ValueType.Double)
            return true;

        if (cellType == CellRawValueType.Numeric && value.ValueType == System.ValueType.Int)
            return true;

        if (cellType == CellRawValueType.DateOnly && value.ValueType == System.ValueType.DateOnly)
            return true;

        if (cellType == CellRawValueType.DateTime && value.ValueType == System.ValueType.DateTime)
            return true;

        if (cellType == CellRawValueType.TimeOnly && value.ValueType == System.ValueType.TimeOnly)
            return true;

        // specific cases
        if (cellType == CellRawValueType.DateTime && value.ValueType == System.ValueType.DateOnly)
            return true;
        if (cellType == CellRawValueType.DateOnly && value.ValueType == System.ValueType.DateTime)
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
    public static bool MatchCellTypeAndIfComparison(CellRawValueType cellType, CellRawValueType cellTypeB)
    {
        if (cellType == CellRawValueType.String && cellTypeB == CellRawValueType.String)
            return true;

        if (cellType == CellRawValueType.Numeric && cellTypeB == CellRawValueType.Numeric)
            return true;

        if (cellType == CellRawValueType.DateOnly && cellTypeB == CellRawValueType.DateOnly)
            return true;

        if (cellType == CellRawValueType.DateTime && cellTypeB == CellRawValueType.DateTime)
            return true;

        if (cellType == CellRawValueType.TimeOnly && cellTypeB == CellRawValueType.TimeOnly)
            return true;

        // specific cases
        if (cellType == CellRawValueType.DateTime && cellTypeB == CellRawValueType.DateOnly)
            return true;
        if (cellType == CellRawValueType.DateOnly && cellTypeB == CellRawValueType.DateTime)
            return true;

        return false;
    }
}