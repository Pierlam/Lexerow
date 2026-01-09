using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

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

    NameObject,

    SetVar,

    CloseExcel,

    ExcelFuncCell,

    ColCellFunc,

    ExcelCellAddress,

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

    And,
    Or,

    EndIf,
    EndOnExcel,

    InstrBlank,
    InstrNull,

    //--process
    ProcessSheets,

    ProcessInstrForEachRow,
    ProcessRow,

    //--Object
    ObjectFilenamesSelected,

    ObjectDate,
    ObjectDateTime,
    ObjectTime,
    ObjectExcelFile,

    //--functions
    FuncSelectFiles,

    FuncDate,

    // expression
    MathExpr,

    BoolExpr
}

/// <summary>
/// Return type of the instruction after the execution of it.
/// </summary>
public enum InstrReturnType
{
    Nothing,

    ValueBool,
    ValueInt,
    ValueDouble,
    ValueString,
    ValueDateOnly,
    ValueDateTime,
    ValueTimeOnly,

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
    /// Return type of the instruction after the execution of it.
    /// exp: InstrValue -> depend on the value type, ValueBool,...
    /// </summary>
    public InstrReturnType ReturnType { get; set; } = InstrReturnType.Nothing;

    /// <summary>
    /// Return true if the instruction is already executed.
    /// </summary>
    public bool IsExecuted { get; set; } = false;
}