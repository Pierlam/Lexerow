using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef;

/// <summary>
/// OnSheet SheetNum/SheetName
///    [FirstRow <val>] [FirstCol <val>]
///    	[OnHeader]  todo: To Be Defined
///    ForEach Row
///         instr..
/// </summary>
public class InstrOnSheet : InstrBase
{
    public InstrOnSheet(ScriptToken scriptToken, InstrValue instrValue) : base(scriptToken)
    {
        InstrType = InstrType.OnSheet;
        InstrFirstDataRow= instrValue;
    }

    // OnHeader instructions
    // TODO:

    /// <summary>
    /// sheet number.
    /// base1 for human, need to apply minus 1 to access sheet.
    /// </summary>
    public int SheetNum { get; set; } = 1;

    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// First row
    /// Human readable, base1
    /// </summary>
    //public int FirstRowNum { get; set; } = 2;

    /// <summary>
    /// First data row value or var or fct call.
    /// Human readable, base1
    /// </summary>
    public InstrBase InstrFirstDataRow { get; set; }


    /// <summary>
    /// Human readable, base1
    /// </summary>
    public int FirstColNum { get; set; } = 1;

    /// <summary>
    /// List of instr to apply on each row.
    /// ForEach Row
    /// </summary>
    public List<InstrBase> ListInstrForEachRow { get; set; } = new List<InstrBase>();
}