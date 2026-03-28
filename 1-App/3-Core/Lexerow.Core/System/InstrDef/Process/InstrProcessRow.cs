using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef.Process;

/// <summary>
/// Process a data row. Execute all defined instructions.
/// ForEach Row; IfThen...
/// Used during execution.
/// </summary>
public class InstrProcessRow : InstrBase
{
    public InstrProcessRow(ScriptToken scriptToken, List<InstrBase> listInstrForEachRow) : base(scriptToken)
    {
        InstrType = InstrType.ProcessRow;
        ListInstrForEachRow = listInstrForEachRow;
    }

    /// <summary>
    /// row address, base 1.
    /// Placed before the first one.
    /// </summary>
    public int RowAddr { get; set; } = 0;

    /// <summary>
    /// column address, base 1.
    /// Placed before the first one.
    /// </summary>
    public int ColAddr { get; set; } = 0;

    /// <summary>
    /// List of Instructions to execute on each datarow.
    /// </summary>
    public List<InstrBase> ListInstrForEachRow { get; set; }
}