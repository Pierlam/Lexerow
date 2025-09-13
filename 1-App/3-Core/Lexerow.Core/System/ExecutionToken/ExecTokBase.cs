using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

public enum ExecTokType
{
    Eol,
    OpenBracket,
    CloseBracket,

    ConstValue,

    ObjectName,

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
/// Base of all execution token classes.
/// Build by the syntax analyzer.
/// </summary>
public abstract class ExecTokBase
{
    public ExecTokBase(ScriptToken scriptToken)
    {
        if(scriptToken != null)
            ListScriptToken.Add(scriptToken);
    }

    /// <summary>
    /// attached to one or more script token.
    /// </summary>
    public List<ScriptToken> ListScriptToken { get; set; }=new List<ScriptToken>();

    public ScriptToken FirstScriptToken()
    {
        return ListScriptToken[0];
    }

    public ExecTokType ExecTokType { get; set; }

    /// <summary>
    /// Return type of this function call.
    /// </summary>
    public InstrFunctionReturnType ReturnType { get; set; } = InstrFunctionReturnType.Nothing;

}
