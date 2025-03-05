using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;


/// <summary>
/// One If-Then instruction.
/// </summary>
public class InstrForEachRowIfThenItem
{
    /// <summary>
    /// If condition instruction.
    /// Only comparison instruction is allowed.
    /// </summary>
    public InstrBase InstrIf { get; set; }

    /// <summary>
    /// Then list of instructions.
    /// only InstrSetCellVal is allowed.
    /// </summary>
    public List<InstrBase> ListInstrThen { get; set; } = new List<InstrBase>();

}

/// <summary>
/// Instruction ForEach Row If Then
/// </summary>
public class InstrForEachRowIfThen : InstrBase
{
    /// <summary>
    /// Constructor 1.
    /// </summary>
    /// <param name="excelFileObjectName"></param>
    /// <param name="sheetNum"></param>
    /// <param name="firstDataRowNum"></param>
    /// <param name="instrIf"></param>
    /// <param name="listInstrThen"></param>
    public InstrForEachRowIfThen(string excelFileObjectName, int sheetNum, int firstDataRowNum, InstrBase instrIf, List<InstrBase> listInstrThen)
    { 
        InstrType = InstrType.ForEachRowIfThen;
        ExcelFileObjectName = excelFileObjectName;
        SheetNum = sheetNum;
        FirstDataRowNum= firstDataRowNum;

        AddInstrIfThen(instrIf, listInstrThen);
    }

    /// <summary>
    /// Constructor 2.
    /// </summary>
    /// <param name="excelFileObjectName"></param>
    /// <param name="sheetNum"></param>
    /// <param name="firstDataRowNum"></param>
    /// <param name="instrIf"></param>
    /// <param name="instrThen"></param>
    public InstrForEachRowIfThen(string excelFileObjectName, int sheetNum, int firstDataRowNum, InstrBase instrIf, InstrBase instrThen)
    {
        InstrType = InstrType.ForEachRowIfThen;
        ExcelFileObjectName = excelFileObjectName;
        SheetNum = sheetNum;
        FirstDataRowNum = firstDataRowNum;

        AddInstrIfThen(instrIf, instrThen);
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
    /// </summary>
    public List<InstrForEachRowIfThenItem> ListInstrIfThen { get; private set; } = new List<InstrForEachRowIfThenItem>();

    /// <summary>
    /// If condition instruction.
    /// Only comparison instruction is allowed.
    /// </summary>
    //public InstrBase InstrIf { get; set; }

    /// <summary>
    /// Then list of instructions.
    /// only InstrSetCellVal is allowed.
    /// </summary>
    //public List<InstrBase> ListInstrThen { get; set; } = new List<InstrBase>();

    public bool AddInstrIfThen(InstrBase instrIf, InstrBase instrThen)
    {
        List<InstrBase> listInstrThen = new List<InstrBase>
        {
            instrThen
        };

        return AddInstrIfThen(instrIf, listInstrThen);

    }

    public bool AddInstrIfThen(InstrBase instrIf, List<InstrBase> listInstrThen)
    {
        // create the first IfThen instr
        InstrForEachRowIfThenItem item = new InstrForEachRowIfThenItem();

        item.InstrIf = instrIf;
        item.ListInstrThen.AddRange(listInstrThen);

        ListInstrIfThen.Add(item);
        return true;
    }
}
