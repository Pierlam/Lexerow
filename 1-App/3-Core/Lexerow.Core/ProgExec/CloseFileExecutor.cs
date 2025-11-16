using Lexerow.Core.System;
using Lexerow.Core.System.Excel;

namespace Lexerow.Core.ProgExec;

internal class CloseFileExecutor
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