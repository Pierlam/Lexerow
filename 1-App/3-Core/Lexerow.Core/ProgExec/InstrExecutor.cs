using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;

namespace Lexerow.Core.ProgExec;

/// <summary>
/// Instruction executor.
/// </summary>
public class InstrExecutor
{
    private IActivityLogger _logger;

    private InstrSelectFilesExecutor _instrSelectFilesExecutor;

    private InstrOnExcelExecutor _instrOnExcelExecutor;

    private InstrOnSheetExecutor _instrOnSheetExecutor;

    private InstrRowExecutor _instrRowExecutor;

    private InstrIfThenElseExecutor _instrIfThenElseExecutor;

    private InstrComparisonExecutor _instrComparisonExecutor;

    private InstrSetColCellFuncExecutor _instrSetColCellFuncExecutor;

    private InstrSetVarExecutor _instrSetVarExecutor;

    public InstrExecutor(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _instrSelectFilesExecutor = new InstrSelectFilesExecutor(_logger);
        _instrOnExcelExecutor = new InstrOnExcelExecutor(_logger, excelProcessor);
        _instrOnSheetExecutor= new InstrOnSheetExecutor(_logger, excelProcessor);
        _instrRowExecutor = new InstrRowExecutor(_logger, excelProcessor);
        _instrIfThenElseExecutor = new InstrIfThenElseExecutor(_logger);
        _instrComparisonExecutor = new InstrComparisonExecutor(_logger, excelProcessor);
        _instrSetColCellFuncExecutor = new InstrSetColCellFuncExecutor(_logger, excelProcessor);
        _instrSetVarExecutor = new InstrSetVarExecutor(_logger, _instrSetColCellFuncExecutor);
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
    public bool ExecInstr(Result result, ProgExecVarMgr progExecVarMgr, InstrBase instr)
    {
        ProgExecContext ctx = new ProgExecContext();

        bool res = true;
        ctx.StackInstr.Push(instr);

        while (true)
        {
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

            if (instr.InstrType == InstrType.SelectFiles)
            {
                res = _instrSelectFilesExecutor.Exec(result, ctx, progExecVarMgr, instr as InstrSelectFiles);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.OnExcel)
            {
                res = _instrOnExcelExecutor.ExecInstrOnExcel(result, ctx, progExecVarMgr, instr as InstrOnExcel);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.ProcessSheets)
            {
                res = _instrOnSheetExecutor.ExecInstrProcessSheets(result, ctx, progExecVarMgr, instr as InstrProcessSheets);
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
                res = _instrComparisonExecutor.ExecInstrComparison(result, ctx, progExecVarMgr, instr as InstrComparison);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.Then)
            {
                res = _instrIfThenElseExecutor.ExecInstrThen(result, ctx, progExecVarMgr, instr as InstrThen);
                if (!res) return false;
                continue;
            }

            // string concatenation, exp: "data" + ".xlsx"
            // TODO:

            result.AddError(ErrorCode.ExecInstrNotManaged, instr.FirstScriptToken());
            return false;
        }
    }
}