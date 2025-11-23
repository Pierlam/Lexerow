using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgExec;
public class InstrRowExecutor
{
    private IActivityLogger _logger;

    private IExcelProcessor _excelProcessor;

    public InstrRowExecutor(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
    }

    /// <summary>
    /// process the current datarow.
    ///     Execute all defined instructions (ForEach Row ListInstr).
    ///  -Stack: ProcessRow, OnSheet, OnExcel
    ///  Next: ForEachRow.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrProcessRow"></param>
    /// <returns></returns>
    public bool ExecInstrProcessRow(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrProcessRow instrProcessRow)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrRowExecutor.ExecInstrProcessRow", string.Empty);

        // next data row exists?
        int lastRowNum = _excelProcessor.GetLastRowNum(ctx.ExcelSheet);
        if (instrProcessRow.RowNum > lastRowNum)
        {
            // no more datarow to process, go back to OnSheet instr
            ctx.StackInstr.Pop();
            _logger.LogExecEnd(ActivityLogLevel.Info, "InstrRowExecutor.ExecInstrProcessRow", "No More row");
            return true;
        }

        ctx.RowNum = instrProcessRow.RowNum;
        // prepare the next one
        instrProcessRow.RowNum++;

        // update insights
        result.Insights.NewRowProcessed();

        // next: process all defined instructions on the current row
        InstrProcessInstrForEachRow instrForEachRow = new InstrProcessInstrForEachRow(instrProcessRow.FirstScriptToken(), instrProcessRow.ListInstrForEachRow);

        ctx.StackInstr.Push(instrForEachRow);

        _logger.LogExecEnd(ActivityLogLevel.Info, "InstrRowExecutor.ExecInstrProcessRow", "NextRowNum: " + instrProcessRow.RowNum);
        return true;
    }

    /// <summary>
    /// Execute next instr defined in the ForEach/Next instr block.
    ///  -Stack: ProcessInstrForEachRow, ProcessRow, OnSheet, OnExcel
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrForEachRow"></param>
    /// <returns></returns>
    public bool ExecProcessInstrForEachRow(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrProcessInstrForEachRow instrForEachRow)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrRowExecutor.ExecProcessInstrForEachRow", string.Empty);

        // execute next instr in ForEach Row
        instrForEachRow.InstrToProcessNum++;

        // next instr exists?
        if (instrForEachRow.InstrToProcessNum >= instrForEachRow.ListInstr.Count)
        {
            // no more instr to execute in OnSheet/ForEachRow
            ctx.StackInstr.Pop();
            _logger.LogExecEnd(ActivityLogLevel.Info, "InstrRowExecutor.RunInstrForEachRow", "No More Instr");
            return true;
        }

        // get the next instr to execute
        InstrBase instrBase = instrForEachRow.ListInstr[instrForEachRow.InstrToProcessNum];
        ctx.StackInstr.Push(instrBase);
        _logger.LogExecEnd(ActivityLogLevel.Info, "InstrRowExecutor.RunInstrForEachRow", "InstrToProcessNum: " + instrForEachRow.InstrToProcessNum);
        return true;
    }

}
