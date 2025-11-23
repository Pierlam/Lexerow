using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

public enum InstrType
{
    Eol,
    OpenBracket,
    CloseBracket,
    Comma,
    Dot,
    Plus,
    Minus,
    Mul,
    Div,
    Percent,

    Value,

    ObjectName,

    Var,
    SetVar,

    SelectFiles,
    CloseExcel,

    ExcelFuncCell,

    ColCellFunc,
    ExcelFileObject,

    ExcelCellAddress,

    //SetCellVal,
    //SetCellNull,
    //SetCellBlank,

    SepComparison,
    Comparison,

    CompOperator,
    CompCellVal,
    CompCellValIsNull,

    CompListColCellAnd,

    ForEachRowIfThen,

    OnExcel,
    OnSheet,
    FirstRow,
    ForEach,
    Row,
    ForEachRow,
    Col,
    Cell,
    Next,
    End,
    If,
    Then,
    IfThenElse,

    EndIf,
    EndOnExcel,

    InstrBlank,
    InstrNull,

    ProcessSheets,
    ProcessInstrForEachRow,
    ProcessRow,
}

public enum InstrFunctionReturnType
{
    Nothing,

    ValueBool,
    ValueInt,
    ValueDouble,
    ValueString,

    ExcelFile,
    ExcelSheet,
    ExcelCols,
    ExcelRows,
    ExcelCol,
    ExcelRow,

    ExcelCell,
    ExcelCells,

    ListString
}

/// <summary>
/// Base of all execution token classes.
/// Build by the syntax analyzer.
/// </summary>
public abstract class InstrBase
{
    public InstrBase(ScriptToken scriptToken)
    {
        if (scriptToken != null)
            ListScriptToken.Add(scriptToken);
    }

    /// <summary>
    /// attached to one or more script token.
    /// </summary>
    public List<ScriptToken> ListScriptToken { get; set; } = new List<ScriptToken>();

    public ScriptToken FirstScriptToken()
    {
        return ListScriptToken[0];
    }

    public InstrType InstrType { get; protected set; }

    /// <summary>
    /// Is it a function call?
    /// exp: OpenExcel
    /// </summary>
    public bool IsFunctionCall { get; protected set; } = false;

    /// <summary>
    /// Return type of this function call.
    /// </summary>
    public InstrFunctionReturnType ReturnType { get; set; } = InstrFunctionReturnType.Nothing;
}