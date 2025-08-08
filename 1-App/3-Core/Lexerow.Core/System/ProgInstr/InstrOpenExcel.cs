using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// instruction: OpenExcel("myfile.xlsx")
/// return an InstrExcelFile object.
/// </summary>
public class InstrOpenExcel : InstrBase
{
    public InstrOpenExcel()
    {
        InstrType = InstrType.OpenExcel;
        ReturnType = InstrFunctionReturnType.ExcelFile;
    }

    /// <summary>
    /// function parameter. 
    /// Should be a string const value, exp: 'MyFile.xlsx'
    /// or a varname, exp: OpenExcel(a)
    /// or a fct call, exp: OpenExcel(GetFileName())  TO_DEFINE
    /// or an expression, exp: OpenExcel('MyFile'+ '.xlsx') TO_DEFINE
    /// </summary>
    //public InstrBase Param { get; set; }
    // TODO: remove !
}
