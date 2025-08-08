using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// To build the instruction OnExcelOnSheet
/// </summary>
public class BuildInstrOnExcelOnSheet
{
    /// <summary>
    /// varname -> without quote, exp: OnExcel file
    /// filename, with quotes, exp: OnExcel "myfile.xlsx"
    /// </summary>
    public string FileName {  get; set; } = string.Empty;

    public bool IsFileNameVar { get; set; } = false;

    public int SheetNum { get; set; } = 0;

    public string SheetName { get; set; } = string.Empty;

    public int FirstRow { get; set; } = 0;

    // list of OnSheet -> List of IfThen -> List of then Instr
    // TODO:
}
