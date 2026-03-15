using DocumentFormat.OpenXml.Spreadsheet;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef.FuncCall;
using Lexerow.Core.System.InstrDef.Object;
using Lexerow.Core.Utils;
using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.InstrProgExec.ExecFunc;

/// <summary>
/// Instr CopyHeader executor.
/// e.g.:
/// 1/ basic case: CopyHeader("data.xlsx", "dataRes.xlsx")  
/// 2/ with varname or excelfile: CopyHeader(file, fileRes)  .
/// 3/ 
/// </summary>
public class InstrFuncCopyHeaderExecutor
{
    private IActivityLogger _logger;
    private ExcelProcessor _excelProcessor;

    public InstrFuncCopyHeaderExecutor(IActivityLogger activityLogger, ExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
    }

    /// <summary>
    /// Execute instr CopyHeader(sourceFile, TargetFile)
    /// By default sheetNum is 0, the first one for source and target file.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="progExecVarMgr"></param>
    /// <param name="instrCopyHeader"></param>
    /// <returns></returns>
    public bool ExecFuncCopyHeader(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrFuncCallCopyHeader instrCopyHeader)
    {
        _logger.LogExecStart(ActivityLogLevel.Debug, "InstrFuncCopyHeaderExecutor.ExecFuncCopyHeader", string.Empty);

        // source file instr need to be executed ?
        if (InstrUtils.NeedToBeExecuted(instrCopyHeader.InstrSourceFile))
        {
            ctx.StackInstr.Push(instrCopyHeader.InstrSourceFile);
            return true;
        }

        // target file instr need to be executed ?
        if (InstrUtils.NeedToBeExecuted(instrCopyHeader.InstrTargetFile))
        {
            ctx.StackInstr.Push(instrCopyHeader.InstrTargetFile);
            return true;
        }

        // source sheetname instr need to be executed ?
        if (InstrUtils.NeedToBeExecuted(instrCopyHeader.InstrSourceSheet))
        {
            ctx.StackInstr.Push(instrCopyHeader.InstrSourceSheet);
            return true;
        }

        // target sheetname instr need to be executed ?
        if (InstrUtils.NeedToBeExecuted(instrCopyHeader.InstrTargetSheet))
        {
            ctx.StackInstr.Push(instrCopyHeader.InstrTargetSheet);
            return true;
        }

        // decode sheet number, default is 0
        int sheetNumSource = 0;
        int sheetNumTarget = 0;

        // the source file can be a string, a varname or an excel file object
        if (!InstrUtils.GetFilenameOrExceFileFromInstr(result, progExecVarMgr, instrCopyHeader.InstrSourceFile, out string filenameSource, out ExcelFile excelFileSource))
            return false;
        if(string.IsNullOrEmpty(filenameSource))
        {
            InstrUtils.GetFirstValueFromInstrObjectSelectedFiles(progExecVarMgr, instrCopyHeader.InstrSourceFile, out filenameSource);
        }

        // the target file can be a string, a varname or an excel file object, same for target file. 
        if (!InstrUtils.GetFilenameOrExceFileFromInstr(result, progExecVarMgr, instrCopyHeader.InstrTargetFile, out string filenameTarget, out ExcelFile excelFileTarget))
            return false;

        // define how to manage the close of the source excel file
        bool closeSourceFile = false;

        // define how to manage the close of the target excel file
        bool closeTargetFile = false;

        try
        {
            // open the source excel file 
            if(excelFileSource==null)
            {
                excelFileSource = _excelProcessor.OpenExcelFile(filenameSource);
                closeSourceFile = true;
            }

            // get the source sheet
            ExcelSheet sheetSource = _excelProcessor.GetSheetAt(excelFileSource, sheetNumSource);

            // from A1 to last column of the first row, copy the header cell value from source sheet to target sheet
            ExcelRow row = _excelProcessor.GetRowAt(sheetSource, 1);
            if(row==null)
            {
                var error = result.AddNewError(ErrorCode.ExecExcelSheetEmpty, instrCopyHeader.FirstScriptToken(), filenameSource);
                _logger.LogExecEndError(error, "InstrFuncCopyHeaderExecutor.ExecFuncCopyHeader", string.Empty);
                return false;
            }

            // open the target excel file 
            if (excelFileTarget == null)
            {
                excelFileTarget = _excelProcessor.OpenExcelFile(filenameTarget);
                closeTargetFile = true;
            }

            // get the target sheet
            ExcelSheet sheetTarget = _excelProcessor.GetSheetAt(excelFileTarget, sheetNumTarget);

            // the target sheet should be empty
            int targetlastRowIndex= _excelProcessor.GetLastRowIndex(sheetTarget);
            if(targetlastRowIndex>0)
            {
                var error = result.AddNewError(ErrorCode.ExecExcelSheetNotEmpty, instrCopyHeader.FirstScriptToken(), filenameTarget);
                _logger.LogExecEndError(error, "InstrFuncCopyHeaderExecutor.ExecFuncCopyHeader", string.Empty);
                return false;
            }


            for (int colNum = 1; colNum <= row.Row.Count(); colNum++)
            {
                // get the cell value
                ExcelCell cell = _excelProcessor.GetCellAt(sheetSource,colNum, 1);
                if(cell==null)continue;

                ExcelCellValue cellValue= _excelProcessor.GetCellValue(sheetSource, cell);
                if (cellValue == null || cellValue.IsEmpty) continue;

                if(cellValue.CellType== ExcelCellType.String)
                {
                    _excelProcessor.SetCellValue(sheetTarget, colNum, 1, cellValue.StringValue);
                    continue;
                }

                // TODO: other type?  For header, we expect the value is string, if not string, we also convert to string and copy to target sheet.
            }

            if(closeSourceFile)
                CloseFileExecutor.Exec(result, _excelProcessor, excelFileSource);

            if (closeTargetFile)
                CloseFileExecutor.Exec(result, _excelProcessor, excelFileTarget);

            // remove the instr from the stack
            ctx.StackInstr.Pop();
            ctx.PrevInstrExecuted = null;
            return true;
        }
        catch (Exception ex)
        {
            var error = result.AddNewError(ErrorCode.ExecUnableCopyHeader, instrCopyHeader.FirstScriptToken(), ex.Message);
            _logger.LogExecEndError(error, "InstrFuncCreateExcelExecutor.ExecFuncCopyHeader", string.Empty);
            return false;
        }
    }
}

