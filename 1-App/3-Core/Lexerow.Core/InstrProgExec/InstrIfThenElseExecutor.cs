using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Execute instruction If-Then-Else.
/// </summary>
public class InstrIfThenElseExecutor
{
    private IActivityLogger _logger;

    public InstrIfThenElseExecutor(IActivityLogger activityLogger)
    {
        _logger = activityLogger;
    }

    public bool ExecInstrIfThenElse(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrIfThenElse instrIfThenElse)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrIfThenElseExecutor.ExecInstrIfThenElse", string.Empty);

        // manage execution of the then instr if if instr result is true
        if (ctx.PrevInstrExecuted != null)
        {
            // instr if executed before?
            InstrIf instrIf = ctx.PrevInstrExecuted as InstrIf;
            if (instrIf != null)
            {
                _logger.LogExecEnd(ActivityLogLevel.Info, "InstrIfThenElseExecutor.ExecInstrIfThenElse", "Prev was If, push Then block instr");
                // execute then instr
                instrIfThenElse.InstrThen.ClearRun();
                ctx.StackInstr.Push(instrIfThenElse.InstrThen);
                ctx.PrevInstrExecuted = null;
                return true;
            }

            // instr Then executed before?
            InstrThen instrThen = ctx.PrevInstrExecuted as InstrThen;
            if (instrThen != null)
            {
                // remove the instr IfThenElse to go back to instr ForEachRow
                //-Stack: IfThenElse, ForEachRow, NextRow, OnSheet, OnExcel
                ctx.StackInstr.Pop();
                ctx.PrevInstrExecuted = null;
                return true;
            }
        }

        // execute If part
        ctx.StackInstr.Push(instrIfThenElse.InstrIf);
        _logger.LogExecEnd(ActivityLogLevel.Info, "InstrIfThenElseExecutor.ExecInstrIfThenElse", "Push If instr");
        return true;
    }

    /// <summary>
    /// execute If part.
    /// 1/ Can be a comparison: If operandLeft operator operandRight
    /// 2/ only one Operand:
    ///   2.1/ bool var: If valExists
    ///   2.2/ function call: If fct()
    ///   2.3/ mat expr: (a+12)
    ///
    /// -Stack: InstrComparison, InstrIf, IfThenElse, ForEachRow, NextRow, OnSheet, OnExcel
    ///
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrIf"></param>
    /// <returns></returns>
    public bool ExecInstrIf(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrIf instrIf)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrIfThenElseExecutor.ExecInstrIf", string.Empty);

        //--if instr already executed?
        if (ctx.PrevInstrExecuted != null)
        {
            //-is it a comparison instr?
            var instrComparison = ctx.PrevInstrExecuted as InstrComparison;
            if (instrComparison != null)
            {
                _logger.LogExecStart(ActivityLogLevel.Info, "InstrIfThenElseExecutor.ExecInstrIf", "Prev was If comparison, result: " + instrIf.Result.ToString());
                instrIf.Result = instrComparison.Result;

                // remove the if from the stack
                ctx.StackInstr.Pop();

                // if cond execution return true, so execute then instr
                if (instrIf.Result)
                {
                    // update insights
                    result.Insights.NewIfCondMatch();

                    // the if instr becomes the previous one
                    ctx.PrevInstrExecuted = instrIf;
                    return true;
                }

                // remove the ifThenElse from the stack, go back to the ForEachRow
                ctx.StackInstr.Pop();
                ctx.PrevInstrExecuted = null;
                return true;
            }

            //-is it a bool var?
            // TODO:
            throw new NotImplementedException("InstrIfThenElseExecutor.ExecInstrIf: Bool Var or fct call to implement");

            //-is it a fct call?
            // TODO:
        }

        //--case1: If operandLeft operator operandRight
        if (instrIf.InstrBase.InstrType == InstrType.Comparison)
        {
            ctx.StackInstr.Push(instrIf.InstrBase);
            _logger.LogExecEnd(ActivityLogLevel.Info, "InstrIfThenElseExecutor.ExecInstrIf", "Run If comparison");
            return true;
        }

        //--case2: If Operand
        // TODO:
        throw new NotImplementedException("InstrIfThenElseExecutor.ExecInstrIf: case2: If Operand to implement");
        return true;
    }

    public bool ExecInstrThen(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrThen instrThen)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrIfThenElseExecutor.ExecInstrThen", string.Empty);

        instrThen.RunInstrNum++;
        if (instrThen.RunInstrNum >= instrThen.ListInstr.Count)
        {
            // no more Then instr to execute
            ctx.StackInstr.Pop();
            ctx.PrevInstrExecuted = instrThen;
            return true;
        }

        InstrBase instrBase = instrThen.ListInstr[instrThen.RunInstrNum];
        ctx.StackInstr.Push(instrBase);
        return true;
    }
}