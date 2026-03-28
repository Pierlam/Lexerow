using Lexerow.Core.System;
using OpenExcelSdk;

namespace Lexerow.Core.InstrProgExec;

internal class CloseFileExecutor
{
    /// <summary>
    /// Close an open excel file.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="excelProcessor"></param>
    /// <param name="excelFile"></param>
    /// <returns></returns>
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