using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

public enum InstrOnExcelBuildStage
{
    OnExcel,
    Files,
    OnSheet,
    SheetNum,
    SheetName,
    ForEach,

    // waiting for the first instr after the token Row
    Row,

    // first instr foudn, waiting for next instr or for the Next token
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
    /// General case on file to analyse.
    /// can be a string const value or a variable.
    /// TO_DEL: replaced by instrBase.
    /// </summary>
    //public List<InstrBase> ListFiles { get; private set; } = new List<InstrBase>();

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
    public InstrSelectFiles InstrSelectFiles { get; set; } = null;


    public List<InstrOnSheet> ListSheets { get; private set; }= new List<InstrOnSheet>();

    /// <summary>
    /// Used by the parser to build the instr.
    /// </summary>
    public InstrOnSheet CurrOnSheet { get; set; } = null;

    ///// <summary>
    ///// Init to -1 which is not started.
    ///// Used in execution.
    ///// </summary>
    //public int FileToProcessNum { get; set; } = -1;

    ///// <summary>
    ///// Used in execution.
    ///// </summary>
    //public int SheetToProcessNum { get; set; } = -1;

    ///// <summary>
    ///// Used in execution.
    ///// TODO: pb! ne stocker que les noms de fichiers , les ouvrir au denier momeont, un par un
    ///// </summary>
    //public List<InstrExcelFileObject> ListExcelFiles { get; set; } = new List<InstrExcelFileObject>();

    /// <summary>
    /// Create a new OnSheet instr, becomes the default one.
    /// </summary>
    /// <param name="scriptToken"></param>
    /// <param name="sheetNum"></param>
    public void CreateOnSheet(ScriptToken scriptToken, int sheetNum)
    {
        InstrOnSheet instrOnSheet = new InstrOnSheet(scriptToken);
        instrOnSheet.SheetNum = sheetNum;
        ListSheets.Add(instrOnSheet);
        CurrOnSheet= instrOnSheet;
    }
}
