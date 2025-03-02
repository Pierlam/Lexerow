using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Utils;
public class ExcelExtendedUtils : ExcelUtils
{
    /// <summary>
    /// Does the cell type match the If-Comparison cell.Value type?
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

}
