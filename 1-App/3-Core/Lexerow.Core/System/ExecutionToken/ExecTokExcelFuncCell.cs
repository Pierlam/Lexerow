using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class ExecTokExcelFuncCell : ExecTokBase
{
    public ExecTokExcelFuncCell(ScriptToken scriptToken):base(scriptToken)
    {
        ExecTokType = ExecTokType.ExcelFuncCell;
    }
}
