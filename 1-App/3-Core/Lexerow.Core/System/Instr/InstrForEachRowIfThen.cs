using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Instruction ForEach Row If Then
/// </summary>
public class InstrForEachRowIfThen : InstrBase
{
    public InstrForEachRowIfThen(string excelFileObjectName, int sheetNum, int firstDataRowNum, InstrBase instrIf, List<InstrBase> listInstrThen)
    { 
        InstrType = InstrType.InstrForEachCellInColsIfThen;
        ExcelFileObjectName = excelFileObjectName;
        SheetNum = sheetNum;
        FirstDataRowNum= firstDataRowNum;
        InstrIf = instrIf;
        ListInstrThen.AddRange(listInstrThen);
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
    /// wehre start the data row, after the header.
    /// By default the header take the first row of the sheet.
    /// base0
    /// </summary>
    public int FirstDataRowNum { get; set; } = 1;

    /// <summary>
    /// If condition instruction.
    /// Only InstrCompCellVal is allowed for now.
    /// </summary>
    public InstrBase InstrIf { get; set; }

    /// <summary>
    /// Then list of instructions.
    /// only InstrSetCellVal is allowed.
    /// </summary>
    public List<InstrBase> ListInstrThen { get; set; } = new List<InstrBase>();
}
