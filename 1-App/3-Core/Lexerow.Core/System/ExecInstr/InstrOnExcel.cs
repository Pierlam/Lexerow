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
///    OnSheet 0,1
///      ForEach Row
///      If A.Cell > 10 Then A.Cell= 10
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
    /// </summary>
    public List<InstrBase> ListFiles { get; private set; } = new List<InstrBase>();

    public List<InstrOnSheet> ListSheets { get; private set; }= new List<InstrOnSheet>();

    public InstrOnSheet CurrOnSheet { get; set; } = null;

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
