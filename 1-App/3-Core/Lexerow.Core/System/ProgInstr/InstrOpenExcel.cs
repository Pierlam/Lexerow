using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrOpenExcel : InstrBase
{
    public InstrOpenExcel(string excelFileObjectName, string fileName)
    {
        InstrType= InstrType.OpenExcel;
        ExcelFileObjectName= excelFileObjectName;
        FileName = fileName; 
    }

    public string ExcelFileObjectName { get; set; }
    public string FileName { get; set; }    
}
