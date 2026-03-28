using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.InstrDef.FuncCall;

/// <summary>
/// instruction: CreateExcel("myfile.xlsx")
/// create a new excel file.
///   CreateExcel("myfile.xlsx")
///   CreateExcel("myfile.xlsx", "sheet1")
/// </summary>
public class InstrFuncCallCreateExcel : InstrBase
{
    public InstrFuncCallCreateExcel(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.FuncCreateExcel;
        IsFunctionCall = true;
        // return an excel file object
        ReturnType = InstrReturnType.ExcelFile;
    }

    /// <summary>
    /// The filename of the excel file to create. It can be a string or an instruction returning a string (like a variable or a function call).
    /// </summary>
    public InstrBase InstrFileName { get; set; }

    /// <summary>
    /// If provided, the sheet will be created with the name. Otherwise, the default name will be used.
    /// </summary>
    public InstrBase? InstrSheetName { get; set; } = null;

    public override string ToString()
    {
        return "CreateExcel()";
    }

}
