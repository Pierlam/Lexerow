using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class ExecTokExcelCellAddress : ExecTokBase
{
    public ExecTokExcelCellAddress(ScriptToken scriptToken, string strCellAddress, int intCellAddress): base(scriptToken)
    {
        ExecTokType= ExecTokType.ExcelCellAddress;
        StrCellAddress= strCellAddress;
        IntCellAddress = intCellAddress;
    }

    public string StrCellAddress { get; set; }
    public  int IntCellAddress { get; set; }
}
