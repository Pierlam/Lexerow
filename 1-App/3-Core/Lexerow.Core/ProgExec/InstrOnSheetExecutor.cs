using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.Utils;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgExec;

/// <summary>
/// Instruction OnSheet Executor.
/// in OnExcel
/// </summary>
internal class InstrOnSheetExecutor
{
    private IActivityLogger _logger;

    private IExcelProcessor _excelProcessor;

    public InstrOnSheetExecutor(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
    }

    /// <summary>
    /// after OnExcel, comes here, process all sheets, one by one.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrNextSheet"></param>
    /// <returns></returns>
    public bool ExecInstrProcessSheets(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrProcessSheets instrNextSheet)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrOnExcelExecutor.ExecInstrProcessSheets", string.Empty);

        // move to next sheet num
        instrNextSheet.SheetNum++;

        if (instrNextSheet.SheetNum >= instrNextSheet.ListSheet.Count)
        {
            // no more sheet to process, got back to the instr OnExcel
            ctx.StackInstr.Pop();
            return true;
        }

        // update insights
        result.Insights.StartNewSheet(instrNextSheet.SheetNum);

        // focus on the current sheet
        InstrOnSheet instrOnSheet = instrNextSheet.ListSheet[instrNextSheet.SheetNum];
        ctx.StackInstr.Push(instrOnSheet);
        ctx.ExcelSheet = null;
        return true;
    }

    /// <summary>
    /// Process a sheet, execute all defined instr in ForEach Row block.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrOnSheet"></param>
    /// <returns></returns>
    public bool ExecInstrOnSheet(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrOnSheet instrOnSheet)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrOnExcelExecutor.ExecInstrOnSheet", string.Empty);
        bool res;

        // start of the sheet processing?
        if (ctx.ExcelSheet == null)
        {
            // get the sheet from excel
            ctx.ExcelSheet = _excelProcessor.GetSheetAt(ctx.ExcelFileObject.ExcelFile, instrOnSheet.SheetNum - 1);

            // process datarow of the current sheet
            InstrProcessRow instrProcessRow = new InstrProcessRow(instrOnSheet.FirstScriptToken(), instrOnSheet.ListInstrForEachRow);

            // get the FirstRow value, can be an instrValue, a var or a fct call.
            if(!GetFirstRowValue(result, ctx, progRunVarMgr, instrOnSheet, out int val))
                return false;

            // translate in base0 from human readable base1
            instrProcessRow.RowNum = val - 1;
            instrProcessRow.ColNum = instrOnSheet.FirstColNum - 1;

            // set to the first datarow
            ctx.StackInstr.Push(instrProcessRow);
            return true;
        }

        // go back to the instr Next Sheet
        ctx.StackInstr.Pop();
        return true;
    }

    /// <summary>
    /// get FirstRow value
    /// can be a InstrValue, InstrObjectName (var) of a fct call.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="progRunVarMgr"></param>
    /// <param name="instrOnSheet"></param>
    /// <param name="instrProcessRow"></param>
    /// <param name="val"></param>
    /// <returns></returns>
    private bool GetFirstRowValue(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrOnSheet instrOnSheet, out int val)
    {
        val = 1;

        //--is it a Value?
        InstrValue instrValue = instrOnSheet.InstrFirstDataRow as InstrValue;
        if (instrValue != null)
        {
            if (!InstrUtils.GetValueIntFromInstrValue(instrValue, instrOnSheet.InstrFirstDataRow.FirstScriptToken().LineNum, out ResultError error, out val))
            {
                result.AddError(error);
                return false;
            }
            // check the int value, should be >= 1
            if (val < 1)
            {
                result.AddError(ErrorCode.ExecValueIntWrong, instrValue.FirstScriptToken());
                return false;
            }

            return true;
        }

        //--is it a Var (ObjectName) ?
        InstrObjectName instrObjectName = instrOnSheet.InstrFirstDataRow as InstrObjectName;
        if (instrObjectName != null) 
        {
            ProgExecVar progExecVar = progRunVarMgr.FindLastInnerVarByName(instrObjectName.ObjectName);
            if(progExecVar == null)
            {
                result.AddError(ErrorCode.ExecInstrVarNotFound, instrValue.FirstScriptToken());
                return false;
            }

            // the value of the value should be an Int value
            if (!InstrUtils.GetValueIntFromInstrValue(progExecVar.Value, instrOnSheet.InstrFirstDataRow.FirstScriptToken().LineNum, out ResultError error, out val))
            {
                result.AddError(error);
                return false;
            }
            // check the int value, should be >= 1
            if (val < 1)
            {
                result.AddError(ErrorCode.ExecValueIntWrong, instrValue.FirstScriptToken());
                return false;
            }

            return true;
        }

        //--is it a fct call?
        // TODO:

        result.AddError(ErrorCode.ExecInstrNotManaged, instrOnSheet.InstrFirstDataRow.FirstScriptToken());
        return false;
    }
}
