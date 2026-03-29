using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
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

public class InstrFuncCopyRowExecutor
{
    private IActivityLogger _logger;
    private ExcelProcessor _excelProcessor;

    public InstrFuncCopyRowExecutor(IActivityLogger activityLogger, ExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
    }

    /// <summary>
    /// Execute the function CopyRow(fileTarget) which copy the current datarow in the fileTarget.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="progExecVarMgr"></param>
    /// <param name="instrCopyRow"></param>
    /// <returns></returns>
    /// <exception cref="NotImplementedException"></exception>
    public bool ExecFuncCopyRow(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrFuncCallCopyRow instrCopyRow)
    {
        _logger.LogExecStart(ActivityLogLevel.Debug, "InstrFuncCopyHeaderExecutor.ExecFuncCopyRow", string.Empty);

        // source file instr need to be executed ?
        if (InstrUtils.NeedToBeExecuted(instrCopyRow.InstrTargetFile))
        {
            ctx.StackInstr.Push(instrCopyRow.InstrTargetFile);
            return true;
        }

        // get the target file
        if (!InstrUtils.GetObjectExcelFileFromInstrVar(result, progExecVarMgr, instrCopyRow.InstrTargetFile, out InstrObjectExcelFile instrObjectExcelFile))
            return false;

        // by default, copy the datarow source to the first sheet
        ExcelSheet excelSheetTarget = _excelProcessor.GetFirstSheet(instrObjectExcelFile.ExcelFile);

        // on source excel file
        int lastColAddr = _excelProcessor.GetLastColAddress(ctx.ExcelSheet, ctx.RowAddr);

        // on target excel file 
        int lastRowIndex = _excelProcessor.GetLastRowIndex(excelSheetTarget);
        lastRowIndex++;

        for (int c = 1; c <= lastColAddr; c++)
        {
            // get the source cell, from the current sheet and datarow
            ExcelCell cell = _excelProcessor.GetCellAt(ctx.ExcelSheet, c, ctx.RowAddr);

            // copy to the target excel file, at the last index
            _excelProcessor.CopyCellValue(ctx.ExcelSheet, cell, excelSheetTarget, ExcelCellAddressUtils.ConvertAddress(c, lastRowIndex));
        }

        // remove the instr from the stack
        ctx.StackInstr.Pop();
        ctx.PrevInstrExecuted = null;
        return true;
    }
}
