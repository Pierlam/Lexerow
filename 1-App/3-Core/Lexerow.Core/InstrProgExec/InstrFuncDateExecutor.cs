using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.FuncCall;
using Lexerow.Core.System.InstrDef.Object;
using Lexerow.Core.Utils;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.InstrProgExec;
public class InstrFuncDateExecutor
{
    IActivityLogger _activityLogger;

    public InstrFuncDateExecutor(IActivityLogger activityLogger)
    {
        _activityLogger = activityLogger;
    }

    /// <summary>
    /// Execute the function Date.
    /// exp: Date(year, month, day)
    /// Generate a date object.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="progExecVarMgr"></param>
    /// <param name="instrFuncCallDate"></param>
    /// <returns></returns>
    public bool ExecFuncDate(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrFuncCallDate instrFuncCallDate)
    {
        // the year param is not a value or a var?
        if (InstrUtils.NeedToBeExecuted(instrFuncCallDate.InstrYear))
        {
            ctx.StackInstr.Push(instrFuncCallDate.InstrYear);
            return true;
        }
        // same for month 
        if (InstrUtils.NeedToBeExecuted(instrFuncCallDate.InstrMonth))
        {
            ctx.StackInstr.Push(instrFuncCallDate.InstrMonth);
            return true;
        }
        // same for day
        if (InstrUtils.NeedToBeExecuted(instrFuncCallDate.InstrDay))
        {
            ctx.StackInstr.Push(instrFuncCallDate.InstrDay);
            return true;
        }

        // get the year int value 
        if (!InstrUtils.GetIntFromInstrExec(result, progExecVarMgr, instrFuncCallDate.InstrYear, out bool _, out int year))
            return false;

        // get the month int value 
        if (!InstrUtils.GetIntFromInstrExec(result, progExecVarMgr, instrFuncCallDate.InstrMonth, out bool _, out int month))
            return false;

        // get the day int value 
        if (!InstrUtils.GetIntFromInstrExec(result, progExecVarMgr, instrFuncCallDate.InstrDay, out bool _, out int day))
            return false;

        // create a dateOnly object 
        try
        {
            DateOnly dateOnly = new DateOnly(year, month, day);
            ValueDateOnly valueDateOnly=new ValueDateOnly(dateOnly);
            InstrValue instrValue = new InstrValue(instrFuncCallDate.FirstScriptToken(), dateOnly);
            // remove the instr FuncDate from the stack
            ctx.StackInstr.Pop();
            ctx.PrevInstrExecuted= instrValue;
            return true;
        }
        catch (Exception ex) 
        {
            result.AddError(ErrorCode.ExecValueDateWrong, instrFuncCallDate.FirstScriptToken());
            return false;
        }

        return false;
    }
}
