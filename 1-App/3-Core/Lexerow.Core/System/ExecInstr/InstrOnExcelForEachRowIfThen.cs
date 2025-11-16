using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;


/// <summary>
/// One If Col Then instruction.
/// </summary>
public class InstrIfColThen : InstrBase
{
    public InstrIfColThen(ScriptToken scriptToken) : base(scriptToken)
    {
    }

    /// <summary>
    /// If condition instruction.
    /// Only comparison instruction is allowed.
    /// </summary>
    public InstrRetBoolBase InstrIf { get; set; }

    /// <summary>
    /// Then list of instructions.
    /// only InstrSetCellVal is allowed.
    /// </summary>
    public List<InstrBase> ListInstrThen { get; set; } = new List<InstrBase>();

}

/// <summary>
/// Instruction ForEach Row If Then
/// </summary>
public class InstrOnExcelForEachRowIfThen : InstrBase
{
    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="excelFileObjectName"></param>
    /// <param name="sheetNum"></param>
    /// <param name="firstDataRowNum"></param>
    /// <param name="instrIf"></param>
    /// <param name="listInstrThen"></param>
    public InstrOnExcelForEachRowIfThen(ScriptToken scriptToken, string excelFileObjectName, int sheetNum, int firstDataRowNum, List<InstrIfColThen> listInstrIfColThen):base(scriptToken)
    { 
        InstrType = InstrType.ForEachRowIfThen;
        ExcelFileObjectName = excelFileObjectName;
        SheetNum = sheetNum;
        FirstDataRowNum= firstDataRowNum;
        ListInstrIfColThen.AddRange(listInstrIfColThen);
    }

    /// <summary>
    /// exp: file
    /// </summary>
    public string ExcelFileObjectName { get; set; } 

    /// <summary>
    /// Excel sheet number.
    /// By default it's the first one: 0.
    /// </summary>
    public int SheetNum { get; set; }

    /// <summary>
    /// Where start the data row, after the header.
    /// By default the header take the first row of the sheet.
    /// base0
    /// </summary>
    public int FirstDataRowNum { get; set; } = 1;

    /// <summary>
    /// List of instr If-Then for the same: excel file, excel sheet, and FirstDataRowNum.
    /// InstrIfColThen
    /// </summary>
    public List<InstrIfColThen> ListInstrIfColThen { get; private set; } = new List<InstrIfColThen>();

}
