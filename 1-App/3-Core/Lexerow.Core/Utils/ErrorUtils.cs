using Lexerow.Core.System;
using OpenExcelSdk.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Utils;

public class ErrorUtils
{
    /// <summary>
    ///  Convert the excel error to a result error
    /// </summary>
    /// <param name="excelError"></param>
    /// <returns></returns>
    public static ResultError Convert(ExcelError excelError)
    {
        ErrorCode errorCode = ErrorCode.ExcelError;

        if(excelError.ErrorCode==  ExcelErrorCode.FileNotFound)
            errorCode = ErrorCode.ExcelUnableOpenFile;
        if (excelError.ErrorCode == ExcelErrorCode.FileAlreadyExists)
            errorCode = ErrorCode.ExcelUnableOpenFile;
        if (excelError.ErrorCode == ExcelErrorCode.FileNull)
            errorCode = ErrorCode.ExcelUnableOpenFile;
        if (excelError.ErrorCode == ExcelErrorCode.UnableOpenFile)
            errorCode = ErrorCode.ExcelUnableOpenFile;
        if (excelError.ErrorCode == ExcelErrorCode.UnableCloseFile)
            errorCode = ErrorCode.ExcelUnableCloseFile;
        if (excelError.ErrorCode == ExcelErrorCode.UnableCreateFile)
            errorCode = ErrorCode.ExcelUnableOpenFile;

        return new ResultError(errorCode, excelError.Exception, excelError.Message);
    }

}
