using Lexerow.Core.ProgExec;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;

/// <summary>
/// InstrSetVar executor.
/// exp:
/// a=12
/// A.Cell=12
/// A.cell=blank
/// </summary>
public class InstrSetVarExecutor
{
    private IActivityLogger _logger;
    private InstrSetColCellFuncExecutor _instrSetColCellFuncRunner;

    public InstrSetVarExecutor(IActivityLogger logger, InstrSetColCellFuncExecutor instrSetColCellFuncRunner)
    {
        _logger = logger;
        _instrSetColCellFuncRunner = instrSetColCellFuncRunner;
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
    /// <param name="instrSetVar"></param>
    /// <returns></returns>
    public bool Exec(ExecResult execResult, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrSetVar instrSetVar)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrSetVarExecutor.Exec", "Token: " + instrSetVar.FirstScriptToken());

        //--case A.Cell= xxx ?
        InstrColCellFunc instrColCellFunc = instrSetVar.InstrLeft as InstrColCellFunc;
        if (instrColCellFunc != null)
            return ExecSetToColCellFunc(execResult, ctx, progExecVarMgr, instrSetVar, instrColCellFunc);

        // the left part should be an objectName (varname), exp: a=xx
        InstrObjectName instrObjectName = instrSetVar.InstrLeft as InstrObjectName;
        if (instrObjectName == null)
        {
            execResult.AddError(ErrorCode.ExecInstrVarTypeNotExpected, "Instr Left: " + instrSetVar.InstrLeft.FirstScriptToken());
            return false;
        }

        //--case a=12, the right instr is a const value
        InstrValue instrConstValue = instrSetVar.InstrRight as InstrValue;
        if (instrConstValue != null)
        {
            // get or create the var, set the value
            CreateVar(ctx, progExecVarMgr, instrSetVar.InstrLeft, instrConstValue);
            ctx.StackInstr.Pop();
            return true;
        }

        //--case a=b, the right instr is a a var too
        instrObjectName = instrSetVar.InstrRight as InstrObjectName;
        if (instrObjectName != null)
        {
            // get the var on the right to catch the result
            //ici();

            // get or create the var, set the value
            CreateVar(ctx, progExecVarMgr, instrSetVar.InstrLeft, instrObjectName);
            ctx.StackInstr.Pop();
            return true;
        }

        // first execute the right part of the SetVar instr
        if (ctx.PrevInstrExecuted == null)
        {
            ctx.StackInstr.Push(instrSetVar.InstrRight);
            return true;
        }

        // get or create the var, set the value, exp: f=SelectFiles()
        CreateVar(ctx, progExecVarMgr, instrSetVar.InstrLeft, ctx.PrevInstrExecuted);

        // now remove the SetInstr from the stack
        ctx.PrevInstrExecuted = null;
        ctx.StackInstr.Pop();
        return true;
    }

    /// <summary>
    /// Execute instr SetVar, left part format is: A.Cell
    /// A.Cell= 12, "hello", blank, null
    /// a.Cell=var
    /// A.Cell=B.Cell
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="progRunVarMgr"></param>
    /// <param name="instrSetVar"></param>
    /// <param name="instrColCellFunc"></param>
    /// <returns></returns>
    private bool ExecSetToColCellFunc(ExecResult execResult, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrSetVar instrSetVar, InstrColCellFunc instrColCellFunc)
    {
        _logger.LogExecOnGoing(ActivityLogLevel.Info, "InstrSetVarExecutor.ExecSetToColCellFunc", "Left is InstrColCellFunc: " + instrSetVar.FirstScriptToken());

        //--case A.cell=10 ?
        InstrValue instrConstValue = instrSetVar.InstrRight as InstrValue;
        if (instrConstValue != null)
        {
            if (!_instrSetColCellFuncRunner.ExecSetCellValue(execResult, ctx.ExcelSheet, ctx.RowNum, instrColCellFunc, instrConstValue))
                return false;
            ctx.StackInstr.Pop();
            return true;
        }

        //--case A.cell=varName ?
        // TODO:

        //--case A.cell=blank ?
        InstrBlank instrBlank = instrSetVar.InstrRight as InstrBlank;
        if (instrBlank != null)
        {
            if (!_instrSetColCellFuncRunner.ExecSetCellBlank(execResult, ctx.ExcelSheet, ctx.RowNum, instrColCellFunc))
                return false;
            ctx.StackInstr.Pop();
            return true;
        }

        //--case A.cell=null ?
        InstrNull instrNull = instrSetVar.InstrRight as InstrNull;
        if (instrNull != null)
        {
            if (!_instrSetColCellFuncRunner.ExecSetCellNull(execResult, ctx.ExcelSheet, ctx.RowNum, instrColCellFunc))
                return false;
            ctx.StackInstr.Pop();
            return true;
        }

        //--case A.Cell= Fct() ?
        // TODO:

        execResult.AddError(ErrorCode.ExecInstrNotManaged, "Instr Right: " + instrSetVar.InstrRight.FirstScriptToken());
        return false;
    }

    private bool CreateVar(ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrBase instrName, InstrBase instrtValue)
    {
        // the var already defined ?
        ProgExecVar execVar = progRunVarMgr.ListExecVar.FirstOrDefault(v => v.AreSame(instrName));

        if (execVar == null)
        {
            // create the var
            execVar = new ProgExecVar(instrName, instrtValue);
            progRunVarMgr.Add(execVar);
        }
        else
            execVar.Value = instrtValue;

        return true;
    }
}