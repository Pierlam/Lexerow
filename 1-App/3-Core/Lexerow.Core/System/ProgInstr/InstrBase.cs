using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

public enum InstrType
{
    Eol,
    OpenBracket,
    CloseBracket,

    ConstValue,

    Var,
    SetVar,

    OpenExcel,
    CloseExcel,

    ExcelFuncCell,

    ExcelFileObject,

    ExcelCellAddress,

    SetCellVal,
    SetCellNull,
    SetCellBlank,

    CompOperator,
    CompCellVal,
    CompCellValIsNull,

    CompListColCellAnd,

    ForEachRowIfThen,

    IfThenElse
}

public enum InstrFunctionReturnType
{
    Nothing,
    ExcelFile,
    ExcelSheet,
    ExcelCols,
    ExcelRows,
    ExcelCol,
    ExcelRow,

    ExcelCell,
    ExcelCells,
}

/// <summary>
/// Base of all instruction.
/// </summary>
public abstract class InstrBase
{
    public InstrType InstrType { get; set; }

    /// <summary>
    /// Return type of this function call.
    /// </summary>
    public InstrFunctionReturnType ReturnType { get; set; } = InstrFunctionReturnType.Nothing;

}
