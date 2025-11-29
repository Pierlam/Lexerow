using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef.Object;

/// <summary>
/// Represent a excel file instance.
/// opened or closed.
/// </summary>
public class InstrObjectExcelFile : InstrBase
{
    public InstrObjectExcelFile(ScriptToken scriptToken, string filename, IExcelFile excelFile) : base(scriptToken)
    {
        InstrType = InstrType.ObjectExcelFile;
        Filename = filename;
        ExcelFile = excelFile;
    }

    public InstrObjectExcelFile(ScriptToken scriptToken, string filename) : base(scriptToken)
    {
        InstrType = InstrType.ObjectExcelFile;
        Filename = filename;
    }

    public string Filename { get; set; }

    /// <summary>
    ///  the excel file object.
    ///  Can be loaded at the last time, just before it will be used.
    /// </summary>
    public IExcelFile? ExcelFile { get; set; } = null;
}