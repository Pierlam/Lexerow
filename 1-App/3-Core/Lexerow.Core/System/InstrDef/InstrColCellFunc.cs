using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

public enum InstrColCellFuncType
{
    Value,
    BgColor,
    FgColor
}

/// <summary>
/// Instruction Column Cell function.
/// Funcrion can be: Value by default, BgColor, FgColor.
/// </summary>
public class InstrColCellFunc : InstrBase
{
    public InstrColCellFunc(ScriptToken scriptToken, InstrColCellFuncType type, string colName, int colNum) : base(scriptToken)
    {
        InstrType = InstrType.ColCellFunc;
        InstrColCellFuncType = InstrColCellFuncType.Value;
        ColName = colName;
        ColNum = colNum;
    }

    public InstrColCellFuncType InstrColCellFuncType { get; private set; }

    /// <summary>
    /// exp: A, DF
    /// </summary>
    public string ColName { get; set; }

    /// <summary>
    /// Exp: 1 for A
    /// 2 for B,...
    /// </summary>
    public int ColNum { get; set; }
}