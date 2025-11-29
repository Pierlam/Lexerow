using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef.Func;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Utils;

namespace Lexerow.Core.System.InstrDef;

public enum InstrOnExcelBuildStage
{
    OnExcel,
    Files,
    OnSheet,
    SheetNum,
    SheetName,
    FirstRow,
    FirstRowValue,
    ForEach,

    // waiting for the first instr after the token Row
    Row,

    // first instr found, waiting for next instr or for the Next token
    RowNext,

    If,

    // token end of the instr: End OnExcel
    EndOnExcel_End,

    // final stage
    EndOfExcel,
}

/// <summary>
/// Main instruction model: OnExcel
/// Exp:
/// OnExcel "data.xlsx"
///    OnSheet 1,2
///      ForEach Row
///        If A.Cell > 10 Then A.Cell= 10
///      Next
///  End OnExcel
/// </summary>
public class InstrOnExcel : InstrBase
{
    public InstrOnExcel(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.OnExcel;
    }

    /// <summary>
    /// Instruction compilation build stages.
    /// </summary>
    public InstrOnExcelBuildStage BuildStage { get; set; } = InstrOnExcelBuildStage.OnExcel;

    /// <summary>
    /// instr, coming  from script.
    /// Convert to the InstrSelectFiles during the execution.
    /// can be:
    ///   1/ a string, exp: OnExcel "data.xlsx"
    ///   2/ a varname (ObjectName), exp: OnExcel filename
    ///      varname can be: 2.1/ a string, 2.2/ an InstrSelectFiles, 2.3/ or a varname (to a string or an InstrSelectFiles).
    ///   3/ a function call (ObjectName), exp: OnExcel GetFiles()
    ///   4/ an InstrSelectFiles, exp: OnExcel "f*.xlsx", +"c*.xlsx", -"file.xlsx"
    ///   5/ an InstrSelectFiles, exp: OnExcel SelectFiles("f*.xlsx", +"c*.xlsx", -"file.xlsx")
    /// </summary>
    public InstrBase? InstrFiles { get; set; } = null;

    /// <summary>
    /// Used by the runner.
    /// the file name (or var or instr selectFiles) is processed and converted to an instrSelectFiles,
    /// specially the member ListFinalFilename, which the final list of files to process.
    /// </summary>
    public InstrFuncSelectFiles InstrSelectFiles { get; set; } = null;

    public List<InstrOnSheet> ListSheets { get; private set; } = new List<InstrOnSheet>();

    /// <summary>
    /// Used by the parser to build the instr.
    /// </summary>
    public InstrOnSheet CurrOnSheet { get; set; } = null;

    /// <summary>
    /// Create a new OnSheet instr, becomes the default one.
    /// </summary>
    /// <param name="scriptToken"></param>
    /// <param name="sheetNum"></param>
    public void CreateOnSheet(ScriptToken scriptToken, int sheetNum)
    {
        InstrValue value = InstrUtils.CreateInstrValueInt(CoreInstr.FirstDataRowNum);

        InstrOnSheet instrOnSheet = new InstrOnSheet(scriptToken, value);
        instrOnSheet.SheetNum = sheetNum;
        ListSheets.Add(instrOnSheet);
        CurrOnSheet = instrOnSheet;
    }
}