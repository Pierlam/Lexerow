using Lexerow.Core.System;
using OpenExcelSdk;

namespace Lexerow.Core.InstrProgExec;

internal class CloseFileExecutor
{
    public static bool Exec(Result result, ExcelProcessor excelProcessor, ExcelFile excelFile)
    {
        try
        {
            excelProcessor.CloseExcelFile(excelFile);
            return true;
        }
        catch (Exception ex)
        {
            result.AddError(ErrorCode.ExecUnableCloseFile, ex, excelFile.Filename);
            return false;
        }
    }
}