using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Represent a excel file instance.
/// opened or closed.
/// </summary>
public class InstrExcelFileObject : InstrBase
{
    public InstrExcelFileObject(ScriptToken scriptToken, string filename, IExcelFile excelFile):base(scriptToken)
    {
        InstrType = InstrType.ExcelFileObject;
        Filename = filename;
        ExcelFile= excelFile;
    }

    public string Filename { get; set; }
    public IExcelFile ExcelFile { get; set; }   
}
