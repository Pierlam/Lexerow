using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

/// <summary>
/// TODO: Really needed?
/// </summary>
public abstract class InstrCallMemberBase : InstrBase
{
    protected InstrCallMemberBase(ScriptToken scriptToken) : base(scriptToken)
    {
    }

    /// <summary>
    /// for InstrExcelFile, can be: FirstSheet, GetSheet[idx], GetSheet[name], ...
    /// </summary>
    public InstrCallMemberBase InstrCallMember { get; set; }
}