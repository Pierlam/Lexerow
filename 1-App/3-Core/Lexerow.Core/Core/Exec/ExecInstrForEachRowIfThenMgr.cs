using Lexerow.Core.System.Excel;
using Lexerow.Core.System;
using Lexerow.Core.System.Exec.Event;
using Lexerow.Core.Utils;
namespace Lexerow.Core;

/// <summary>
/// Execute instruciotn: ForEachRow IfThen
/// </summary>
public class ExecInstrForEachRowIfThenMgr
{
    static int _dataRowCount = 0;

    /// <summary>
    /// If Condition Fired count
    /// TODO: plutot IFcondition qui est vrai!
    /// </summary>
    static int _ifConditionFiredCount = 0;

    //static Action<InstrBaseExecEvent> _eventOccurs;
    static Action<AppTrace> _appTraceEvent;

    /// <summary>
    /// Execute the instr on each cols, on each datarow.
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instr"></param>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    public static ExecResult Exec(Action<AppTrace> appTraceEvent, DateTime execStart, IExcelProcessor excelProcessor, IExcelFile excelFile, InstrForEachRowIfThen instr)
    {
        _appTraceEvent = appTraceEvent;
        ExecResult execResult = new ExecResult();
        _dataRowCount = 0;
        _ifConditionFiredCount = 0;

        // only one root/starting sheet
        IExcelSheet sheet = excelProcessor.GetSheetAt(excelFile, instr.SheetNum);
        if (sheet == null)
        {
            CoreError error= new CoreError(ErrorCode.UnableFindSheetByNum, instr.SheetNum.ToString());
            execResult.AddError(error);

            SendAppTraceExec(AppTraceLevel.Error, "ExecInstrForEachRowIfThenMgr.Exec", InstrForEachRowIfThenExecEvent.CreateFinishedError(execStart, error));
            //FireEvent(InstrForEachRowIfThenExecEvent.CreateFinishedError(execStart, error));
            return execResult;
        }
        int currRowNum = instr.FirstDataRowNum;
        int lastRowNum = excelProcessor.GetLastRowNum(sheet);

        // apply IfThen instructions on each datarow
        while (currRowNum <= lastRowNum)
        {
            foreach(var instrIfThen in instr.ListInstrIfColThen)
            {
                execResult = ExecOnDataRow(execStart, excelProcessor, excelFile, sheet, instrIfThen.InstrIf, instrIfThen.ListInstrThen, currRowNum);
                if(!execResult.Result)
                {
                    CoreError error = new CoreError(ErrorCode.InternalError, instr.SheetNum.ToString());
                    execResult.AddError(error);

                    SendAppTraceExec(AppTraceLevel.Error, "ExecInstrForEachRowIfThenMgr.Exec", InstrForEachRowIfThenExecEvent.CreateFinishedError(execStart, error));
                    //FireEvent(InstrForEachRowIfThenExecEvent.CreateFinishedError(execStart, error));
                    return execResult;
                }
            }
            currRowNum++;
            _dataRowCount++;
        }

        //FireEvent(InstrForEachRowIfThenExecEvent.CreateFinishedOk(execStart, _dataRowCount, _ifConditionFiredCount));
        SendAppTraceExec(AppTraceLevel.Info, "ExecInstrForEachRowIfThenMgr.Exec", InstrForEachRowIfThenExecEvent.CreateFinishedOk(execStart, _dataRowCount, _ifConditionFiredCount));

        return execResult;
    }

    /// <summary>
    /// Execute one If-Then instructions on the current datarow.
    /// </summary>
    /// <param name="execStart"></param>
    /// <param name="excelProcessor"></param>
    /// <param name="excelFile"></param>
    /// <param name="excelSheet"></param>
    /// <param name="instrIf"></param>
    /// <param name="listInstrThen"></param>
    /// <param name="rowNum"></param>
    /// <returns></returns>
    static ExecResult ExecOnDataRow(DateTime execStart, IExcelProcessor excelProcessor, IExcelFile excelFile, IExcelSheet excelSheet, InstrRetBoolBase instrIf, List<InstrBase> listInstrThen, int rowNum)
    {
        ExecResult execResult; 

        // execute If cond on the cells of the datarow
        execResult= ExecIfCondition(excelProcessor, excelFile, excelSheet, instrIf, rowNum, out bool condResult); 
        if(!execResult.Result)
        {
            SendAppTraceExec(AppTraceLevel.Error, "ExecInstrForEachRowIfThenMgr.ExecOnDataRow", InstrForEachRowIfThenExecEvent.CreateFinishedError(execStart, execResult.ListError.FirstOrDefault()));
            //FireEvent(InstrForEachRowIfThenExecEvent.CreateFinishedError(execStart, execResult.ListError.FirstOrDefault()));
            // error occurs during the If condition instr execution
            return execResult;
        }

        // the If condition return false, so doesn't execute Then instructions
        if(!condResult)
            return execResult;

        _ifConditionFiredCount++;
        SendAppTraceExec(AppTraceLevel.Info, "ExecInstrForEachRowIfThenMgr.ExecOnDataRow", InstrForEachRowIfThenExecEvent.CreateFinishedInProgress(execStart, _dataRowCount, _ifConditionFiredCount));
        //FireEvent(InstrForEachRowIfThenExecEvent.CreateFinishedInProgress(execStart, _dataRowCount, _ifConditionFiredCount));

        // execute Then instructions
        foreach (InstrBase instrThen in listInstrThen)
        {
            // TODO:  manage error
            execResult = ExecOnCellInstrThen(excelProcessor, excelSheet, instrThen, rowNum);
        }

        return execResult;
    }

    /// <summary>
    /// Execute If cond on the cells of the datarow.
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="excelFile"></param>
    /// <param name="excelSheet"></param>
    /// <param name="instrIf"></param>
    /// <param name="rowNum"></param>
    /// <param name="condResult"></param>
    /// <returns></returns>
    static ExecResult ExecIfCondition(IExcelProcessor excelProcessor, IExcelFile excelFile, IExcelSheet excelSheet, InstrRetBoolBase instrIf, int rowNum, out bool condResult)
    {
        ExecResult execResult= new ExecResult();
        condResult = false;

        //--is the If cond instr InstrCompCellVal or InstrCompCellValIsNull?
        InstrCompColCellVal instrCompCellVal = instrIf as InstrCompColCellVal;
        if (instrCompCellVal != null)
        {
            return ExecInstrCompColCellVal(excelProcessor, excelFile, excelSheet, instrCompCellVal, rowNum, out condResult);
        }


        //--is the If cond instr InstrCompCellValIsNull?
        InstrCompColCellValIsNull instrCompCellValIsNull = instrIf as InstrCompColCellValIsNull;
        if(instrCompCellValIsNull!=null)
        {
            return ExecInstrCompColCellValIsNull(excelProcessor, excelFile, excelSheet, instrCompCellValIsNull, rowNum, out condResult);
        }

        //--is the If cond instr InstrCompListColCellAnd?
        InstrCompListColCellAnd instrCompListColCellAnd = instrIf as InstrCompListColCellAnd;
        if(instrCompListColCellAnd!=null)
        {
            return ExecInstrCompListColCellAnd(excelProcessor, excelFile, excelSheet, instrCompListColCellAnd, rowNum, out condResult);
        }

        //--Not allowed
        execResult.AddError(new CoreError(ErrorCode.IfConditionInstrNotAllowed, instrIf.InstrType.ToString()));
        return execResult;
    }

    static ExecResult ExecInstrCompColCellVal(IExcelProcessor excelProcessor, IExcelFile excelFile, IExcelSheet excelSheet, InstrCompColCellVal instrCompColCellVal, int rowNum, out bool condResult)
    {
        ExecResult execResult = new ExecResult();
        condResult = false;

        var cell = excelProcessor.GetCellAt(excelSheet, rowNum, instrCompColCellVal.ColNum);

        // get the cell value type            
        CellRawValueType cellType = excelProcessor.GetCellValueType(excelSheet, cell);

        // does the cell type match the If-Comparison cell.Value type?
        if (!ExcelExtendedUtils.MatchCellTypeAndIfComparison(cellType, instrCompColCellVal.Value))
            return execResult;

        // execute the If part: comparison condition
        return ExecInstrCompMgr.ExecInstrCompCellVal(excelProcessor, instrCompColCellVal, cell, out condResult);
    }

    static ExecResult ExecInstrCompColCellValIsNull(IExcelProcessor excelProcessor, IExcelFile excelFile, IExcelSheet excelSheet, InstrCompColCellValIsNull instrCompColCellValIsNull, int rowNum, out bool condResult)
    {
        var cell = excelProcessor.GetCellAt(excelSheet, rowNum, instrCompColCellValIsNull.ColNum);

        // execute the If part: comparison condition
        return ExecInstrCompMgr.ExecInstrCompCellValIsNull(excelProcessor, instrCompColCellValIsNull, cell, out condResult);
    }

    static ExecResult ExecInstrCompListColCellAnd(IExcelProcessor excelProcessor, IExcelFile excelFile, IExcelSheet excelSheet, InstrCompListColCellAnd instrCompListColCellAnd, int rowNum, out bool condResult)
    {
        ExecResult execResult = new ExecResult();
        bool condResultLocal;
        condResult = true;

        foreach (InstrRetBoolBase instrRetBoolBase in instrCompListColCellAnd.ListInstrComp)
        {
            // stop when the condition result becomes false
            if (!condResult)
                break;

            //--is the If cond instr InstrCompCellVal or InstrCompCellValIsNull?
            InstrCompColCellVal instrCompCellVal = instrRetBoolBase as InstrCompColCellVal;
            if (instrCompCellVal != null)
            {
                var execResultLocal= ExecInstrCompColCellVal(excelProcessor, excelFile, excelSheet, instrCompCellVal, rowNum, out condResultLocal);
                if (!execResultLocal.Result)
                    execResult.AddListError(execResultLocal.ListError);
                condResult &= condResultLocal;
                continue;
            }

            InstrCompColCellValIsNull instrCompCellValIsNull = instrRetBoolBase as InstrCompColCellValIsNull;
            if (instrCompCellValIsNull != null)
            {
                var execResultLocal = ExecInstrCompColCellValIsNull(excelProcessor, excelFile, excelSheet, instrCompCellValIsNull, rowNum, out condResultLocal);
                if (!execResultLocal.Result)
                    execResult.AddListError(execResultLocal.ListError);
                condResult &= condResultLocal;
            }
        }

        return execResult;
    }

    /// <summary>
    /// Execute the then instruction.
    /// For now, only InstrSetCellVal is authorized.
    /// exp: Cell.Val= 12
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instr"></param>
    /// <param name="sheet"></param>
    /// <param name="colNum"></param>
    /// <param name="rowNum"></param>
    /// <returns></returns>
    static ExecResult ExecOnCellInstrThen(IExcelProcessor excelProcessor, IExcelSheet sheet, InstrBase instr, int rowNum)
    {
        ExecResult execResult= new ExecResult();    

        // check allowed Then instructions

        InstrSetCellVal instrSetCellVal = instr as InstrSetCellVal;
        if(instrSetCellVal!=null)
            return  ExecInstrSetCellMgr.ExecSetCellVal(excelProcessor, instrSetCellVal, sheet, rowNum);

        InstrSetCellNull instrSetCellNull = instr as InstrSetCellNull;
        if(instrSetCellNull!=null)
            return ExecInstrSetCellMgr.ExecSetCellNull(excelProcessor, instrSetCellNull, sheet, rowNum);

        InstrSetCellValueBlank instrSetCellValueBlank = instr as InstrSetCellValueBlank;
        if (instrSetCellValueBlank != null)
            return ExecInstrSetCellMgr.ExecSetCellValueBlank(excelProcessor, instrSetCellValueBlank, sheet, rowNum);

        execResult.AddError(new CoreError(ErrorCode.ThenConditionInstrNotAllowed, instr.ToString()));
        return execResult;
    }

    //static void FireEvent(InstrBaseExecEvent execEvent)
    //{
    //    if (_eventOccurs != null)
    //        _eventOccurs(execEvent);
    //}

    public static void SendAppTraceExec(AppTraceLevel level, string msg, InstrBaseExecEvent execEvent)
    {
        if (_appTraceEvent == null) return;

        AppTrace appTrace = new AppTrace(AppTraceModule.Exec, level, msg, execEvent);
        _appTraceEvent(appTrace);
    }

}
