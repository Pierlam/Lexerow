using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

/// <summary>
/// varName= instr
/// can be a const value, int, string, ..
/// or the result of an instruction.
/// exp: file= OpenExcel('list.xlsx')
/// </summary>
public class InstrSetVar : InstrBase
{
    public InstrSetVar(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.SetVar;
    }

    /// <summary>
    /// The left part of the SetVar instruction.
    /// before the char equal.
    /// Used if the left part is an object.
    /// exp:
    ///   file=
    ///   A.Cell=
    ///   A.Cell.BgColor=
    /// </summary>
    public InstrBase? InstrLeft { get; set; } = null;

    /// <summary>
    /// the right part, after the equal
    /// =12        -> InstrValue
    /// =a         -> InstrVar
    /// =sheet.Cell(A,1)  -> InstrExcelValue
    /// </summary>
    public InstrBase InstrRight { get; set; }

    public override string ToString()
    {
        string instrLeftStr = "(null)";
        if (InstrLeft != null) instrLeftStr = InstrLeft.ToString();
        string instrRightStr = "(null)";
        if (InstrRight != null) instrRightStr = InstrRight.ToString();

        return "SetVar: " + instrLeftStr + "= " + instrRightStr;
    }
}