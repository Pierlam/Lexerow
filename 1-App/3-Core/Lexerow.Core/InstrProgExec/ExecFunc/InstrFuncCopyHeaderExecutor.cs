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
        _logger.Log(ActivityLogLevel.Debug, "InstrFuncCopyHeaderExecutor.ExecFuncCopyHeader", string.Empty);

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
            ExcelSheet excelSheetSource = _excelProcessor.GetSheetAt(excelFileSource, sheetNumSource);

            // header is always on the first row
            int rowAddr = 1;

            // on source excel file
            int lastColAddr = _excelProcessor.GetLastColAddress(excelSheetSource, rowAddr);
            if (lastColAddr == 0)
            {
                var error = result.AddNewError(ErrorCode.ExecExcelSheetEmpty, instrCopyHeader.FirstScriptToken(), filenameSource);
                _logger.LogError("InstrFuncCopyHeaderExecutor.ExecFuncCopyHeader", error);
                return false;
            }

            // open the target excel file 
            if (excelFileTarget == null)
            {
                excelFileTarget = _excelProcessor.OpenExcelFile(filenameTarget);
                closeTargetFile = true;
            }

            // get the target sheet
            ExcelSheet excelSheetTarget = _excelProcessor.GetSheetAt(excelFileTarget, sheetNumTarget);

            // the target sheet should be empty
            int targetlastRowIndex= _excelProcessor.GetLastRowIndex(excelSheetTarget);
            if(targetlastRowIndex>0)
            {
                var error = result.AddNewError(ErrorCode.ExecExcelSheetNotEmpty, instrCopyHeader.FirstScriptToken(), filenameTarget);
                _logger.LogError("InstrFuncCopyHeaderExecutor.ExecFuncCopyHeader", error);
                return false;
            }

            // copy cells of the source row header to the target
            for (int c = 1; c <= lastColAddr; c++)
            {
                // get the source cell
                ExcelCell cellSource = _excelProcessor.GetCellAt(excelSheetSource, c, rowAddr);

                // copy to the target excel file
                _excelProcessor.CopyCellValue(excelSheetSource, cellSource, excelSheetTarget, ExcelCellAddressUtils.ConvertAddress(c, rowAddr));
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
            _logger.LogError("InstrFuncCreateExcelExecutor.ExecFuncCopyHeader", error);
            return false;
        }
    }
}

