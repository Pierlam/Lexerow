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
    /// row index, base 1.
    /// Placed before the first one.
    /// </summary>
    public int RowIndex { get; set; } = 0;

    /// <summary>
    /// column index, base 1.
    /// Placed before the first one.
    /// </summary>
    public int ColIndex { get; set; } = 0;

    /// <summary>
    /// List of Instructions to execute on each datarow.
    /// </summary>
    public List<InstrBase> ListInstrForEachRow { get; set; }
}