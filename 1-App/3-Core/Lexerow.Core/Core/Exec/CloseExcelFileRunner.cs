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
    public static bool Exec(ExecResult execResult, IExcelProcessor excelProcessor, IExcelFile excelFile)
    {
        excelProcessor.Save(excelFile);

        if (!excelProcessor.Close(excelFile, out var error))
        {
            execResult.ListError.Add(error);
            return false;
        }
        return true;
    }

}
