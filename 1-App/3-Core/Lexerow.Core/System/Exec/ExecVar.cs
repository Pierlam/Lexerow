using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

public enum ExecVarType
{
    // string, int, double, ...
    BasicValue,

    ExcelFile

}

public class ExecVar
{
    public ExecVar(string name, ExecVarType execVarType, InstrBase value)
    {
        Name = name;
        ExecVarType = execVarType;
        Value = value;
    }

    public string Name { get; set; }
    
    public ExecVarType ExecVarType { get; set; }

    public InstrBase Value { get; set; }
}
