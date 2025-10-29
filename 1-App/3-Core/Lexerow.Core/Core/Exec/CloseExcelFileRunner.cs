using Lexerow.Core.System;
using Lexerow.Core.System.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Core.Exec;

/// <summary>
/// Keep it!!
/// </summary>
public class CloseExcelFileRunner
{
    public static bool Exec(IExcelProcessor excelProcessor, IExcelFile excelFile, out ExecResultError error)
    {
        excelProcessor.Save(excelFile);

        return excelProcessor.Close(excelFile, out error);
    }

}
