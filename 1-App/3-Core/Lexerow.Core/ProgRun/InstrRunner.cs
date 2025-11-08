using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.ProgRun;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgRun;

/// <summary>
/// Instruction runner.
/// </summary>
public class InstrRunner
{
    IActivityLogger _logger;

    InstrSelectFilesRunner _instrSelectFilesRunner;

    InstrOnExcelRunner _instrOnExcelRunner;

    InstrIfThenElseRunner _instrIfThenElseRunner;

    InstrComparisonRunner _instrComparisonRunner;

    InstrSetColCellFuncRunner _instrSetColCellFuncRunner;

    InstrSetVarRunner _instrSetVarRunner;

    public InstrRunner(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _instrSelectFilesRunner = new InstrSelectFilesRunner(_logger, excelProcessor);
        _instrOnExcelRunner = new InstrOnExcelRunner(_logger, excelProcessor);
        _instrIfThenElseRunner = new InstrIfThenElseRunner(_logger);
        _instrComparisonRunner= new InstrComparisonRunner(_logger, excelProcessor);
        _instrSetColCellFuncRunner = new InstrSetColCellFuncRunner(_logger, excelProcessor);
        _instrSetVarRunner =new InstrSetVarRunner(_logger, _instrSetColCellFuncRunner);
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
    public bool ExecInstr(ExecResult execResult, ProgRunVarMgr progRunVarMgr, InstrBase instr)
    {
        ProgramRunnerContext ctx = new ProgramRunnerContext();

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
                res = _instrSetVarRunner.Run(execResult, ctx, progRunVarMgr, instr as InstrSetVar);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.SelectFiles)
            {
                res = _instrSelectFilesRunner.Run(execResult, ctx, progRunVarMgr, instr as InstrSelectFiles);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.OnExcel)
            {
                res = _instrOnExcelRunner.RunInstrOnExcel(execResult, ctx, progRunVarMgr, instr as InstrOnExcel);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.NextSheet)
            {
                res = _instrOnExcelRunner.RunInstrNextSheet(execResult, ctx, progRunVarMgr, instr as InstrNextSheet);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.OnSheet)
            {
                res = _instrOnExcelRunner.RunInstrOnSheet(execResult, ctx, progRunVarMgr, instr as InstrOnSheet);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.NextRow)
            {
                res = _instrOnExcelRunner.RunInstrNextRow(execResult, ctx, progRunVarMgr, instr as InstrNextRow);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.ForEachRow)
            {
                res = _instrOnExcelRunner.RunInstrForEachRow(execResult, ctx, progRunVarMgr, instr as InstrForEachRow);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.IfThenElse)
            {
                res = _instrIfThenElseRunner.RunInstrIfThenElse(execResult, ctx, progRunVarMgr, instr as InstrIfThenElse);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.If)
            {
                res = _instrIfThenElseRunner.RunInstrIf(execResult, ctx, progRunVarMgr, instr as InstrIf);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.Comparison)
            {
                res = _instrComparisonRunner.RunInstrComparison(execResult, ctx, progRunVarMgr, instr as InstrComparison);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.Then)
            {
                res = _instrIfThenElseRunner.RunInstrThen(execResult, ctx, progRunVarMgr, instr as InstrThen);
                if (!res) return false;
                continue;
            }

            // string concatenation, exp: "data" + ".xlsx"
            // TODO:

            execResult.AddError(ErrorCode.RunInstrNotManaged, instr.FirstScriptToken());
            return false;
        }
    }


}
