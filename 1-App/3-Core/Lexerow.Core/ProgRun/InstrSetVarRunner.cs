using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ProgRun;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgRun;
public class InstrSetVarRunner
{
    IActivityLogger _logger;
    InstrSetColCellFuncRunner _instrSetColCellFuncRunner;

    public InstrSetVarRunner(IActivityLogger logger, InstrSetColCellFuncRunner instrSetColCellFuncRunner)
    {
        _logger = logger;
        _instrSetColCellFuncRunner= instrSetColCellFuncRunner;
    }

    /// <summary>
    /// Execute instr SetVar.
    /// A.Cell=12
    /// A.Cell=b
    /// a=12
    /// a=b
    /// file=SelectFiles(..)
    /// in these cases, create a var name and set the value.
    /// 
    /// special case: A.Cell=xx  
    ///   Not really a set var instruction.
    ///   >Just set the value to the excel cell, doesn't create a var.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instr"></param>
    /// <returns></returns>
    public bool Run(ExecResult execResult, ProgramRunnerContext ctx, ProgRunVarMgr progRunVarMgr, InstrSetVar instr)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "InstrSetVarRunner.Run", "Token: " + instr.FirstScriptToken());

        //--case A.Cell= xxx ?
        InstrColCellFunc instrColCellFunc = instr.InstrLeft as InstrColCellFunc;
        if (instrColCellFunc != null)
        {
            _logger.LogRunOnGoing(ActivityLogLevel.Info, "InstrSetVarRunner.Run", "Left is InstrColCellFunc: " + instr.FirstScriptToken());

            // case A.cell=10 ?
            InstrConstValue instrConstValue1 = instr.InstrRight as InstrConstValue;
            if (instrConstValue1 != null)
            {
                if (!_instrSetColCellFuncRunner.RunSetCellValue(execResult, ctx.ExcelSheet, ctx.RowNum, instrColCellFunc, instrConstValue1))
                    return false;
                ctx.StackInstr.Pop();
                return true;
            }

            // case A.cell=varName ?
            // TODO:

            // case A.cell=null ?
            // TODO:

            // case A.cell=blank ?
            // TODO:

            //--case A.Cell= Fct() ?
            // TODO:

            execResult.AddError(ErrorCode.RunInstrNotManaged, "Instr Right: " + instr.InstrRight.FirstScriptToken());
            return false;
        }

        // the left part should be an objectName (varname), exp: a=xx
        InstrObjectName instrObjectName = instr.InstrLeft as InstrObjectName;
        if(instrObjectName==null)
        {
            execResult.AddError(ErrorCode.RunInstrVarTypeNotExpected, "Instr Left: " + instr.InstrLeft.FirstScriptToken());
            return false;
        }

        //--case a=12, the right instr is a const value
        InstrConstValue instrConstValue = instr.InstrRight as InstrConstValue;
        if (instrConstValue != null)
        {
            // get or create the var, set the value
            CreateVar(ctx, progRunVarMgr, instr.InstrLeft, instrConstValue);
            ctx.StackInstr.Pop();
            return true;
        }

        //--case a=b, the right instr is a a var too
        instrObjectName = instr.InstrRight as InstrObjectName;
        if (instrObjectName != null) 
        {
            // get or create the var, set the value
            CreateVar(ctx, progRunVarMgr, instr.InstrLeft, instrObjectName);
            ctx.StackInstr.Pop();
            return true;
        }

        // first execute the right part of the SetVar instr
        if (ctx.PrevInstrExecuted == null)
        {
            ctx.StackInstr.Push(instr.InstrRight);
            return true;
        }

        // get or create the var, set the value
        CreateVar(ctx, progRunVarMgr, instr.InstrLeft, ctx.PrevInstrExecuted);

        // now remove the SetInstr from the stack
        ctx.PrevInstrExecuted = null;
        ctx.StackInstr.Pop();
        return true;
    }


    bool CreateVar(ProgramRunnerContext ctx, ProgRunVarMgr progRunVarMgr, InstrBase instrName, InstrBase instrtValue)
    {
        // the var already defined ?
        ProgRunVar execVar = progRunVarMgr.ListRunVar.FirstOrDefault(v => v.AreSame(instrName));

        if (execVar == null)
        {
            // create the var
            execVar = new ProgRunVar(instrName, instrtValue);
            progRunVarMgr.Add(execVar);
        }
        else
            execVar.Value = instrtValue;

        return true;
    }

}
