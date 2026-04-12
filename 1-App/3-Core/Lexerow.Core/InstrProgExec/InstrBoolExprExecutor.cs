using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.Utils;
using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Execute boolean instructions.
/// exp: IF A.Cell>10 And A.Cell<2à
/// </summary>
public class InstrBoolExprExecutor
{
    private IActivityLogger _logger;

    private ExcelProcessor _excelProcessor;

    public InstrBoolExprExecutor(IActivityLogger activityLogger, ExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
    }

    public bool ExecInstrBoolExpr(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrBoolExpr instrBoolExpr)
    {
        _logger.LogExec(ActivityLogLevel.Debug, "InstrBoolExprExecutor.ExecInstrBoolExpr", string.Empty);

        //--a previous instr exists
        //if (ctx.PrevInstrExecuted != null)
        //{ }

        // find the next part to be executed
        InstrBase instr= FindNextToExecute(instrBoolExpr, InstrType.BoolExpr);

        // all sub instr are executed, execute the bool expr
        if (instr == instrBoolExpr)
        {
            ExecuteBoolExpr(result, ctx, progExecVarMgr, instrBoolExpr);
            instrBoolExpr.IsExecuted = true;
            ctx.PrevInstrExecuted = instrBoolExpr;
            // clear all IsExecuted flags
            InstrUtils.ClearIsExecuted(instrBoolExpr);
            ctx.StackInstr.Pop();
            return true;
        }

        if (instr != null) 
        {
            ctx.StackInstr.Push(instr);
            return true;
        }


        return true;
    }

    /// <summary>
    /// Find the next sub instruction to be executed.
    /// </summary>
    /// <param name="instrBoolExpr"></param>
    /// <returns></returns>
    InstrBase FindNextToExecute(InstrBoolExpr instrBoolExpr, InstrType instrType)
    {
        bool isExecutedFinal= true;

        foreach (InstrBase instrBase in instrBoolExpr.ListOperand)
        {

            if(NeedToBeExecuted(instrBase, InstrType.BoolExpr, out bool isExecuted))return instrBase;
            isExecutedFinal = isExecutedFinal && isExecuted;
        }
        if (isExecutedFinal)return instrBoolExpr;
        return null;
    }


    bool NeedToBeExecuted(InstrBase instrBase, InstrType instrType, out bool isExecuted)
    {
        isExecuted = false;

        if (instrBase == null) return false;

        if (instrBase.IsExecuted)
        {
            isExecuted = true;
            return false;
        }

        if (instrBase.IsFunctionCall) return true;


        if (instrType != InstrType.BoolExpr && instrBase.InstrType == InstrType.BoolExpr) return true;
        if (instrType != InstrType.Comparison && instrBase.InstrType == InstrType.Comparison) return true;
        if (instrType != InstrType.MathExpr && instrBase.InstrType == InstrType.MathExpr) return true;
        // no
        return false;
    }

    bool ExecuteBoolExpr(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrBoolExpr instrBoolExpr)
    {
        bool resCurr = false;

        foreach (var instr in instrBoolExpr.ListOperand) 
        {
            InstrComparison instrComparison = (InstrComparison)instr;
            if(instrComparison!=null)
                resCurr = instrComparison.Result;

            // And operator
            if (instrBoolExpr.Operator == InstrBoolExprOperator.And)
            {
                instrBoolExpr.Result = resCurr;
                if (!resCurr)
                    // no need to continue one operand is false, the bool expr is false
                    return true;

            }else
            {
                // Or operator
                instrBoolExpr.Result |= resCurr;
            }

        }

        return true;
    }

}
