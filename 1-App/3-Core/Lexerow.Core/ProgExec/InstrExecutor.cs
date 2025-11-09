using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.ProgRun;

namespace Lexerow.Core.ProgExec;

/// <summary>
/// Instruction executor.
/// </summary>
public class InstrExecutor
{
    IActivityLogger _logger;

    InstrSelectFilesExecutor _instrSelectFilesExecutor;

    InstrOnExcelExecutor _instrOnExcelExecutor;

    InstrIfThenElseExecutor _instrIfThenElseExecutor;

    InstrComparisonExecutor _instrComparisonExecutor;

    InstrSetColCellFuncExecutor _instrSetColCellFuncExecutor;

    InstrSetVarExecutor _instrSetVarExecutor;

    public InstrExecutor(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _instrSelectFilesExecutor = new InstrSelectFilesExecutor(_logger, excelProcessor);
        _instrOnExcelExecutor = new InstrOnExcelExecutor(_logger, excelProcessor);
        _instrIfThenElseExecutor = new InstrIfThenElseExecutor(_logger);
        _instrComparisonExecutor= new InstrComparisonExecutor(_logger, excelProcessor);
        _instrSetColCellFuncExecutor = new InstrSetColCellFuncExecutor(_logger, excelProcessor);
        _instrSetVarExecutor =new InstrSetVarExecutor(_logger, _instrSetColCellFuncExecutor);
    }

    /// <summary>
    /// Execute instruction and dependent instructions, until the stack becomes empty. 
    /// InstrRunner
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="listInstr"></param>
    /// <param name="listVar"></param>
    /// <param name="instr"></param>
    /// <returns></returns>
    public bool ExecInstr(ExecResult execResult, ProgExecVarMgr progExecVarMgr, InstrBase instr)
    {
        ProgramExecContext ctx = new ProgramExecContext();

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
                res = _instrSetVarExecutor.Exec(execResult, ctx, progExecVarMgr, instr as InstrSetVar);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.SelectFiles)
            {
                res = _instrSelectFilesExecutor.Exec(execResult, ctx, progExecVarMgr, instr as InstrSelectFiles);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.OnExcel)
            {
                res = _instrOnExcelExecutor.ExecInstrOnExcel(execResult, ctx, progExecVarMgr, instr as InstrOnExcel);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.ProcessSheets)
            {
                res = _instrOnExcelExecutor.ExecInstrProcessSheets(execResult, ctx, progExecVarMgr, instr as InstrProcessSheets);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.OnSheet)
            {
                res = _instrOnExcelExecutor.ExecInstrOnSheet(execResult, ctx, progExecVarMgr, instr as InstrOnSheet);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.ProcessRow)
            {
                res = _instrOnExcelExecutor.ExecInstrProcessRow(execResult, ctx, progExecVarMgr, instr as InstrProcessRow);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.ProcessInstrForEachRow)
            {
                res = _instrOnExcelExecutor.ExecProcessInstrForEachRow(execResult, ctx, progExecVarMgr, instr as InstrProcessInstrForEachRow);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.IfThenElse)
            {
                res = _instrIfThenElseExecutor.ExecInstrIfThenElse(execResult, ctx, progExecVarMgr, instr as InstrIfThenElse);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.If)
            {
                res = _instrIfThenElseExecutor.ExecInstrIf(execResult, ctx, progExecVarMgr, instr as InstrIf);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.Comparison)
            {
                res = _instrComparisonExecutor.ExecInstrComparison(execResult, ctx, progExecVarMgr, instr as InstrComparison);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.Then)
            {
                res = _instrIfThenElseExecutor.ExecInstrThen(execResult, ctx, progExecVarMgr, instr as InstrThen);
                if (!res) return false;
                continue;
            }

            // string concatenation, exp: "data" + ".xlsx"
            // TODO:

            execResult.AddError(ErrorCode.ExecInstrNotManaged, instr.FirstScriptToken());
            return false;
        }
    }


}
