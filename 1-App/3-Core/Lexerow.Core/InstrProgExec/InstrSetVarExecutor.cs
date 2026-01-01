using Lexerow.Core.InstrProgExec;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.Utils;

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
    private InstrSetColCellFuncExecutor _instrSetColCellFuncExecutor;

    public InstrSetVarExecutor(IActivityLogger logger, InstrSetColCellFuncExecutor instrSetColCellFuncRunner)
    {
        _logger = logger;
        _instrSetColCellFuncExecutor = instrSetColCellFuncRunner;
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
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrSetVar"></param>
    /// <returns></returns>
    public bool Exec(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrSetVar instrSetVar)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrSetVarExecutor.Exec", "Token: " + instrSetVar.FirstScriptToken());

        InstrBase instrRight = instrSetVar.InstrRight;

        //--a previous instr exists
        if (ctx.PrevInstrExecuted != null)
        {
            instrRight = ctx.PrevInstrExecuted;
        }

        // the left operand is a function call, so execute it and come back here
        if (InstrUtils.NeedToBeExecuted(instrRight))
        {
            ctx.StackInstr.Push(instrRight);
            //instrComparison.LastInstrExecuted = 1;
            return true;
        }

        //--is the right instr a varname?
        InstrNameObject instrNameObject = instrRight as InstrNameObject;
        if (instrNameObject != null)
        {
            ProgExecVar execVar = progExecVarMgr.FindLastInnerVarByName(instrNameObject.Name);
            instrRight = execVar.Value;
        }

        //--is it A.Cell= xxx ?
        InstrColCellFunc instrColCellFunc = instrSetVar.InstrLeft as InstrColCellFunc;
        if (instrColCellFunc != null)
            return ExecSetToColCellFunc(result, ctx, instrSetVar, instrColCellFunc, instrRight);

        //--is it a=val/var/object?
        InstrNameObject instrObjectName = instrSetVar.InstrLeft as InstrNameObject;
        if (instrObjectName != null)
            return ExecSetToVar(result, ctx, progExecVarMgr, instrSetVar, instrObjectName, instrRight);

        //--is it Fct()=val/var?
        // TODO:

        // now remove the SetInstr from the stack
        ctx.PrevInstrExecuted = null;
        ctx.StackInstr.Pop();
        return true;

        //-XXXXXXXXXXXXXXXXX

        // the left part should be an objectName (varname), exp: a=xx
        //InstrNameObject instrObjectName = instrSetVar.InstrLeft as InstrNameObject;
        //if (instrObjectName == null)
        //{
        //    result.AddError(ErrorCode.ExecInstrVarTypeNotExpected, "Instr Left: " + instrSetVar.InstrLeft.FirstScriptToken());
        //    return false;
        //}

        //--case a=12, the right instr is a const value
        //InstrValue instrValue = instrSetVar.InstrRight as InstrValue;
        //if (instrValue != null)
        //{
        //    // get or create the var, set the value
        //    CreateVar(ctx, progExecVarMgr, instrSetVar.InstrLeft, instrValue);
        //    ctx.StackInstr.Pop();
        //    return true;
        //}

        //--case a=b, the right instr is a a var too
        instrObjectName = instrSetVar.InstrRight as InstrNameObject;
        if (instrObjectName != null)
        {
            // get the var on the right to catch the result
            //ici();

            // get or create the var, set the value
            CreateVar(result, ctx, progExecVarMgr, instrSetVar.InstrLeft, instrObjectName);
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
        CreateVar(result, ctx, progExecVarMgr, instrSetVar.InstrLeft, ctx.PrevInstrExecuted);

        // now remove the SetInstr from the stack
        ctx.PrevInstrExecuted = null;
        ctx.StackInstr.Pop();
        return true;
    }

    /// <summary>
    /// To REMOVE when the new version will be up!
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="progExecVarMgr"></param>
    /// <param name="instrSetVar"></param>
    /// <returns></returns>
    //public bool Exec_OLD(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrSetVar instrSetVar)
    //{
    //    _logger.LogExecStart(ActivityLogLevel.Info, "InstrSetVarExecutor.Exec", "Token: " + instrSetVar.FirstScriptToken());

    //    //--case A.Cell= xxx ?
    //    InstrColCellFunc instrColCellFunc = instrSetVar.InstrLeft as InstrColCellFunc;
    //    if (instrColCellFunc != null)
    //        return ExecSetToColCellFunc(result, ctx, progExecVarMgr, instrSetVar, instrColCellFunc);

    //    // the left part should be an objectName (varname), exp: a=xx
    //    InstrNameObject instrObjectName = instrSetVar.InstrLeft as InstrNameObject;
    //    if (instrObjectName == null)
    //    {
    //        result.AddError(ErrorCode.ExecInstrVarTypeNotExpected, "Instr Left: " + instrSetVar.InstrLeft.FirstScriptToken());
    //        return false;
    //    }

    //    //--case a=12, the right instr is a const value
    //    InstrValue instrValue = instrSetVar.InstrRight as InstrValue;
    //    if (instrValue != null)
    //    {
    //        // get or create the var, set the value
    //        CreateVar(ctx, progExecVarMgr, instrSetVar.InstrLeft, instrValue);
    //        ctx.StackInstr.Pop();
    //        return true;
    //    }

    //    //--case a=b, the right instr is a a var too
    //    instrObjectName = instrSetVar.InstrRight as InstrNameObject;
    //    if (instrObjectName != null)
    //    {
    //        // get the var on the right to catch the result
    //        //ici();

    //        // get or create the var, set the value
    //        CreateVar(ctx, progExecVarMgr, instrSetVar.InstrLeft, instrObjectName);
    //        ctx.StackInstr.Pop();
    //        return true;
    //    }

    //    // first execute the right part of the SetVar instr
    //    if (ctx.PrevInstrExecuted == null)
    //    {
    //        ctx.StackInstr.Push(instrSetVar.InstrRight);
    //        return true;
    //    }

    //    // get or create the var, set the value, exp: f=SelectFiles()
    //    CreateVar(ctx, progExecVarMgr, instrSetVar.InstrLeft, ctx.PrevInstrExecuted);

    //    // now remove the SetInstr from the stack
    //    ctx.PrevInstrExecuted = null;
    //    ctx.StackInstr.Pop();
    //    return true;
    //}

    /// <summary>
    /// Execute instr SetVar, left part format is: A.Cell
    /// A.Cell= 12, "hello", blank, null
    /// a.Cell=var
    /// A.Cell=B.Cell
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="progRunVarMgr"></param>
    /// <param name="instrSetVar"></param>
    /// <param name="instrColCellFunc"></param>
    /// <returns></returns>
    private bool ExecSetToColCellFunc(Result result, ProgExecContext ctx, InstrSetVar instrSetVar, InstrColCellFunc instrColCellFunc, InstrBase instrRight)
    {
        _logger.LogExecOnGoing(ActivityLogLevel.Info, "InstrSetVarExecutor.ExecSetToColCellFunc", "Left is InstrColCellFunc: " + instrSetVar.FirstScriptToken());

        //--case A.cell=10 ?
        InstrValue instrValue = instrRight as InstrValue;
        if (instrValue != null)
        {
            if (!_instrSetColCellFuncExecutor.ExecSetCellValue(result, ctx.ExcelSheet, ctx.RowNum, instrColCellFunc, instrValue))
                return false;
            ctx.StackInstr.Pop();
            return true;
        }

        //--case A.cell=blank ?
        InstrBlank instrBlank = instrRight as InstrBlank;
        if (instrBlank != null)
        {
            if (!_instrSetColCellFuncExecutor.ExecSetCellBlank(result, ctx.ExcelSheet, ctx.RowNum, instrColCellFunc))
                return false;
            ctx.StackInstr.Pop();
            return true;
        }

        //--case A.cell=null ?
        InstrNull instrNull = instrRight as InstrNull;
        if (instrNull != null)
        {
            if (!_instrSetColCellFuncExecutor.ExecSetCellNull(result, ctx.ExcelSheet, ctx.RowNum, instrColCellFunc))
                return false;
            ctx.StackInstr.Pop();
            return true;
        }

        //--case A.Cell= Fct() ?
        // TODO:

        result.AddError(ErrorCode.ExecInstrNotManaged, "Instr Right: " + instrSetVar.InstrRight.FirstScriptToken());
        return false;
    }

    /// <summary>
    /// SetVar a=someting
    /// can be a Value, an object.
    /// Var and fctcall are managed before coming in this function.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="progExecVarMgr"></param>
    /// <param name="instrSetVar"></param>
    /// <param name="instrNameObject"></param>
    /// <param name="instrRight"></param>
    /// <returns></returns>
    private bool ExecSetToVar(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrSetVar instrSetVar, InstrNameObject instrNameObject, InstrBase instrRight)
    {
        // get or create the var, set the value
        CreateVar(result, ctx, progExecVarMgr, instrSetVar.InstrLeft, instrRight);
        ctx.PrevInstrExecuted = null;
        ctx.StackInstr.Pop();
        return true;
    }

    /// <summary>
    /// Create or update a variable, a system one or a user defined one.
    /// </summary>
    /// <param name="ctx"></param>
    /// <param name="progExecVarMgr"></param>
    /// <param name="instrName"></param>
    /// <param name="instrtValue"></param>
    /// <returns></returns>
    private bool CreateVar(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrBase instrName, InstrBase instrBase)
    {
        // is it a system var?
        if (instrName.FirstScriptToken().ScriptTokenType == Lexerow.Core.System.ScriptDef.ScriptTokenType.SystName) 
        {
            // check var name

            // check value
            InstrValue instrValue = instrBase as InstrValue ;
            if(instrValue==null)
            {
                result.AddError(ErrorCode.ExecInstrValueExpected, instrBase.FirstScriptToken());
                return false;
            }

            // it's a system variable
            if (progExecVarMgr.UpdateSystemVar(instrName.FirstScriptToken().Value, instrValue.ValueBase) == null)
            {
                result.AddError(ErrorCode.ExecUnableCreateUpdateVar, instrName.FirstScriptToken());
                return false;
            }

            return true;
        }

        // it's a user defined variable
        if (progExecVarMgr.CreateOrUpdateVar(instrName, instrBase) == null)
        {
            result.AddError(ErrorCode.ExecUnableCreateUpdateVar, instrName.FirstScriptToken());
            return false;
        }

        return true;
    }
}