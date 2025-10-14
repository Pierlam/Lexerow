using Lexerow.Core.System;
using Lexerow.Core.System.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Core.Exec;
public class ExecInstrCloseExcelFileMgr
{
    public static bool Exec(IExcelProcessor excelProcessor, IExcelFile excelFile)
    {
        excelProcessor.Save(excelFile);

        return excelProcessor.Close(excelFile);
    }

}
