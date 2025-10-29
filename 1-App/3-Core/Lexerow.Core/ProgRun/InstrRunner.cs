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

    InstrOpenExcelRunner _instrOpenExcelRunner;

    InstrOnExcelRunner _instrOnExcelRunner;

    InstrIfThenElseRunner _instrIfThenElseRunner;

    InstrComparisonRunner _instrComparisonRunner;

    InstrSetColCellFuncRunner _instrSetColCellFuncRunner;

    public InstrRunner(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _instrOpenExcelRunner = new InstrOpenExcelRunner(_logger, excelProcessor);
        _instrOnExcelRunner = new InstrOnExcelRunner(_logger, excelProcessor);
        _instrIfThenElseRunner = new InstrIfThenElseRunner(_logger);
        _instrComparisonRunner= new InstrComparisonRunner(_logger, excelProcessor);
        _instrSetColCellFuncRunner = new InstrSetColCellFuncRunner(_logger, excelProcessor);
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
    public bool ExecInstr(ExecResult execResult, List<ExecVar> listVar, InstrBase instr)
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
                res = ExecInstrSetVar(execResult, ctx, listVar, instr as InstrSetVar);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.OpenExcel)
            {
                res = _instrOpenExcelRunner.Run(execResult, ctx, listVar, instr as InstrOpenExcel);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.OnExcel)
            {
                res = _instrOnExcelRunner.RunInstrOnExcel(execResult, ctx, listVar, instr as InstrOnExcel);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.NextSheet)
            {
                res = _instrOnExcelRunner.RunInstrNextSheet(execResult, ctx, listVar, instr as InstrNextSheet);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.OnSheet)
            {
                res = _instrOnExcelRunner.RunInstrOnSheet(execResult, ctx, listVar, instr as InstrOnSheet);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.NextRow)
            {
                res = _instrOnExcelRunner.RunInstrNextRow(execResult, ctx, listVar, instr as InstrNextRow);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.ForEachRow)
            {
                res = _instrOnExcelRunner.RunInstrForEachRow(execResult, ctx, listVar, instr as InstrForEachRow);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.IfThenElse)
            {
                res = _instrIfThenElseRunner.RunInstrIfThenElse(execResult, ctx, listVar, instr as InstrIfThenElse);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.If)
            {
                res = _instrIfThenElseRunner.RunInstrIf(execResult, ctx, listVar, instr as InstrIf);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.Comparison)
            {
                res = _instrComparisonRunner.RunInstrComparison(execResult, ctx, listVar, instr as InstrComparison);
                if (!res) return false;
                continue;
            }

            if (instr.InstrType == InstrType.Then)
            {
                res = _instrIfThenElseRunner.RunInstrThen(execResult, ctx, listVar, instr as InstrThen);
                if (!res) return false;
                continue;
            }

            execResult.AddError(ErrorCode.RunInstrNotManaged, instr.FirstScriptToken());
            return false;
        }
    }

    /// <summary>
    /// Execute instr SetVar.
    /// exp: a=12  file=SelectExcel(..)
    /// in these cases, create a var name and set the value.
    /// 
    /// special case: A.Cell= 12
    /// Just set the value to the excel cell, doesn't create a var.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instr"></param>
    /// <returns></returns>
    bool ExecInstrSetVar(ExecResult execResult, ProgramRunnerContext ctx,  List<ExecVar> listVar, InstrSetVar instr)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "ProgRunner.ExecInstrSetVar", "Pushed on stack");

        //--case A.cell= xxx ?
        InstrColCellFunc instrColCellFunc= instr.InstrLeft as InstrColCellFunc;
        if (instrColCellFunc!=null)
        {
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

        }

        // the right instr  is a const value
        InstrConstValue instrConstValue = instr.InstrRight as InstrConstValue;
        if (instrConstValue != null)
        {
            // get or create the var, set the value
            CreateVar(ctx, listVar, instr.InstrLeft, instrConstValue);
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
        CreateVar(ctx, listVar, instr.InstrLeft, ctx.PrevInstrExecuted);

        // now remove the SetInstr from the stack
        ctx.PrevInstrExecuted = null;
        ctx.StackInstr.Pop();
        return true;
    }


    bool CreateVar(ProgramRunnerContext ctx, List<ExecVar> listVar, InstrBase instrName, InstrBase instrtValue)
    {
        // the var already defined ?
        ExecVar execVar = listVar.FirstOrDefault(v => v.AreSame(instrName));

        if (execVar == null)
        {
            // create the var
            execVar = new ExecVar(instrName, instrtValue);
            listVar.Add(execVar);
        }
        else
            execVar.Value = instrtValue;

        return true;
    }

}
