using Lexerow.Core.InstrProgExec.ExecFunc;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.FuncCall;
using Lexerow.Core.System.InstrDef.Process;
using OpenExcelSdk;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Instruction executor.
/// </summary>
public class InstrExecutor
{
    private IActivityLogger _logger;

    private InstrOnExcelExecutor _instrOnExcelExecutor;

    private InstrOnSheetExecutor _instrOnSheetExecutor;

    private InstrRowExecutor _instrRowExecutor;

    private InstrIfThenElseExecutor _instrIfThenElseExecutor;

    private InstrComparisonExecutor _instrComparisonExecutor;

    private InstrBoolExprExecutor _instrBoolExprExecutor;

    private InstrSetColCellFuncExecutor _instrSetColCellFuncExecutor;

    private InstrSetVarExecutor _instrSetVarExecutor;

    private InstrFuncSelectFilesExecutor _instrFuncSelectFilesExecutor;

    private InstrFuncDateExecutor _instrFuncDateExecutor;
    private InstrFuncCreateExcelExecutor _instrFuncCreateExcelExecutor;

    private InstrFuncCopyHeaderExecutor _instrFuncCopyHeaderExecutor;

    private InstrFuncCopyRowExecutor _instrFuncCopyRowExecutor;

    private ProgExecVarMgr _progExecVarMgr;

    /// <summary>
    /// Program instruction executor constructor.
    /// </summary>
    /// <param name="activityLogger"></param>
    /// <param name="excelProcessor"></param>
    /// <param name="progExecVarMgr"></param>
    public InstrExecutor(IActivityLogger activityLogger, ExcelProcessor excelProcessor, ProgExecVarMgr progExecVarMgr)
    {
        _logger = activityLogger;
        _progExecVarMgr = progExecVarMgr;

        _instrOnExcelExecutor = new InstrOnExcelExecutor(_logger, excelProcessor);
        _instrOnSheetExecutor = new InstrOnSheetExecutor(_logger, excelProcessor);
        _instrRowExecutor = new InstrRowExecutor(_logger, excelProcessor);
        _instrIfThenElseExecutor = new InstrIfThenElseExecutor(_logger);
        _instrComparisonExecutor = new InstrComparisonExecutor(_logger, excelProcessor);
        _instrBoolExprExecutor= new InstrBoolExprExecutor(_logger, excelProcessor);
        _instrSetColCellFuncExecutor = new InstrSetColCellFuncExecutor(_logger, excelProcessor, progExecVarMgr);
        _instrSetVarExecutor = new InstrSetVarExecutor(_logger, _instrSetColCellFuncExecutor);

        _instrFuncSelectFilesExecutor = new InstrFuncSelectFilesExecutor(_logger);
        _instrFuncDateExecutor = new InstrFuncDateExecutor(_logger);
        _instrFuncCreateExcelExecutor= new InstrFuncCreateExcelExecutor(_logger, excelProcessor);
        _instrFuncCopyHeaderExecutor= new InstrFuncCopyHeaderExecutor(_logger, excelProcessor);
        _instrFuncCopyRowExecutor= new InstrFuncCopyRowExecutor(_logger, excelProcessor);
    }

    /// <summary>
    /// Execute instruction and dependent instructions, until the stack becomes empty.
    /// InstrRunner
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listInstr"></param>
    /// <param name="listVar"></param>
    /// <param name="instr"></param>
    /// <returns></returns>
    public bool ExecInstr(Result result, Program program, ProgExecVarMgr progExecVarMgr, InstrBase instr)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrExecutor.ExecInstr", instr.ToString());

        ProgExecContext ctx = new ProgExecContext();

        bool res = true;
        ctx.StackInstr.Push(instr);

        while (true)
        {
            // no more instr to execute, exit
            if (ctx.StackInstr.Count == 0)
            {
                return res;
            }

            // get the top instr of the stack
            instr = ctx.StackInstr.Peek();

            if (instr.InstrType == InstrType.SetVar)
            {
                res = _instrSetVarExecutor.Exec(result, ctx, progExecVarMgr, instr as InstrSetVar);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.OnExcel)
            {
                res = _instrOnExcelExecutor.ExecInstrOnExcel(result, ctx, progExecVarMgr, instr as InstrOnExcel);
                _logger.LogExecEnd(ActivityLogLevel.Info, "InstrExecutor.ExecInstr", string.Empty);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.ProcessSheets)
            {
                res = _instrOnSheetExecutor.ExecInstrPerformSheets(result, ctx, progExecVarMgr, instr as InstrProcessSheets);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.OnSheet)
            {
                res = _instrOnSheetExecutor.ExecInstrOnSheet(result, ctx, progExecVarMgr, instr as InstrOnSheet);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.ProcessRow)
            {
                res = _instrRowExecutor.ExecInstrProcessRow(result, ctx, progExecVarMgr, instr as InstrProcessRow);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.ProcessInstrForEachRow)
            {
                res = _instrRowExecutor.ExecProcessInstrForEachRow(result, ctx, progExecVarMgr, instr as InstrProcessInstrForEachRow);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.IfThenElse)
            {
                res = _instrIfThenElseExecutor.ExecInstrIfThenElse(result, ctx, progExecVarMgr, instr as InstrIfThenElse);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.If)
            {
                res = _instrIfThenElseExecutor.ExecInstrIf(result, ctx, progExecVarMgr, instr as InstrIf);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.Comparison)
            {
                res = _instrComparisonExecutor.ExecInstrCompExpr(result, ctx, progExecVarMgr, instr as InstrComparison);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.BoolExpr)
            {
                res = _instrBoolExprExecutor.ExecInstrBoolExpr(result, ctx, progExecVarMgr, instr as InstrBoolExpr);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.Then)
            {
                res = _instrIfThenElseExecutor.ExecInstrThen(result, ctx, progExecVarMgr, instr as InstrThen);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.FuncSelectFiles)
            {
                res = _instrFuncSelectFilesExecutor.Exec(result, ctx, program, instr as InstrFuncCallSelectFiles);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.FuncDate)
            {
                res = _instrFuncDateExecutor.ExecFuncDate(result, ctx, progExecVarMgr, instr as InstrFuncCallDate);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.FuncCreateExcel)
            {
                res = _instrFuncCreateExcelExecutor.ExecFuncCreateExcel(result, ctx, progExecVarMgr, instr as InstrFuncCallCreateExcel);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.FuncCopyHeader)
            {
                res = _instrFuncCopyHeaderExecutor.ExecFuncCopyHeader(result, ctx, progExecVarMgr, instr as InstrFuncCallCopyHeader);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.FuncCopyRow)
            {
                res = _instrFuncCopyRowExecutor.ExecFuncCopyRow(result, ctx, progExecVarMgr, instr as InstrFuncCallCopyRow);
                if (!res) return false;
                continue;
            }

            // string concatenation, exp: "data" + ".xlsx"
            // TODO:

            var error = result.AddNewError(ErrorCode.ExecInstrNotManaged, instr.FirstScriptToken());
            _logger.LogExecEndError(error, "InstrExecutor.ExecInstr", instr.ToString() + " not managed");
            return false;
        }
    }
}