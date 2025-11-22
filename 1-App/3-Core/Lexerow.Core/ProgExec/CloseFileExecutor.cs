using Lexerow.Core.System;
using Lexerow.Core.System.Excel;

namespace Lexerow.Core.ProgExec;

internal class CloseFileExecutor
{
    public static bool Exec(Result result, IExcelProcessor excelProcessor, IExcelFile excelFile)
    {
        excelProcessor.Save(excelFile);

        if (!excelProcessor.Close(excelFile, out var error))
        {
            result.ListError.Add(error);
            return false;
        }
        return true;
    }
}