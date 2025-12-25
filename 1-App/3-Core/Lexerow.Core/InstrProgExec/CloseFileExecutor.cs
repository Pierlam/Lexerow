using Lexerow.Core.System;
using Lexerow.Core.Utils;
using OpenExcelSdk;
using OpenExcelSdk.System;

namespace Lexerow.Core.InstrProgExec;

internal class CloseFileExecutor
{
    public static bool Exec(Result result, ExcelProcessor excelProcessor, ExcelFile excelFile)
    {

        if (!excelProcessor.Close(excelFile, out var error))
        {
            
            result.ListError.Add(ErrorUtils.Convert(error));
            return false;
        }
        return true;
    }
}