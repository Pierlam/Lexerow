using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class ExecTokExcelCellAddress : InstrBase
{
    public ExecTokExcelCellAddress(string strCellAddress, int intCellAddress)
    {
        InstrType= InstrType.ExcelCellAddress;
        StrCellAddress= strCellAddress;
        IntCellAddress = intCellAddress;
    }

    public string StrCellAddress { get; set; }
    public  int IntCellAddress { get; set; }
}
