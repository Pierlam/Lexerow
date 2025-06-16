using Lexerow.Core.System.Excel;
using Lexerow.Core.System;
using Lexerow.Core.System.Exec.Event;
using Lexerow.Core.Utils;
namespace Lexerow.Core;

/// <summary>
/// Execute instruction: ForEachRow IfThen
/// </summary>
public class ExecInstrForEachRowIfThenMgr
{
    static int _dataRowCount = 0;

    /// <summary>
    /// If Condition Fired count
    /// </summary>
    static int _ifConditionFiredCount = 0;

    static Action<AppTrace> _appTraceEvent;

    /// <summary>
    /// Execute the instr on each cols, on each datarow.
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instr"></param>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    public static bool Exec(ExecResult execResult, Action<AppTrace> appTraceEvent, DateTime execStart, IExcelProcessor excelProcessor, IExcelFile excelFile, InstrOnExcelForEachRowIfThen instr)
    {
        _appTraceEvent = appTraceEvent;
        _dataRowCount = 0;
        _ifConditionFiredCount = 0;

        // only one root/starting sheet
        IExcelSheet sheet = excelProcessor.GetSheetAt(excelFile, instr.SheetNum);
        if (sheet == null)
        {
            ExecResultError error= new ExecResultError(ErrorCode.UnableFindSheetByNum, instr.SheetNum.ToString());
            execResult.AddError(error);

            SendAppTraceExec(AppTraceLevel.Error, "ExecInstrForEachRowIfThenMgr.Exec", InstrForEachRowIfThenExecEvent.CreateFinishedError(execStart, error));
            return false;
        }
        int currRowNum = instr.FirstDataRowNum;
        int lastRowNum = excelProcessor.GetLastRowNum(sheet);

        // apply IfThen instructions on each datarow
        while (currRowNum <= lastRowNum)
        {
            foreach(var instrIfThen in instr.ListInstrIfColThen)
            {
                bool res = ExecOnDataRow(execResult, execStart, excelProcessor, excelFile, sheet, instrIfThen.InstrIf, instrIfThen.ListInstrThen, currRowNum);
                if(!res)
                {
                    ExecResultError error = new ExecResultError(ErrorCode.InternalError, instr.SheetNum.ToString());
                    execResult.AddError(error);

                    SendAppTraceExec(AppTraceLevel.Error, "ExecInstrForEachRowIfThenMgr.Exec", InstrForEachRowIfThenExecEvent.CreateFinishedError(execStart, error));
                    return false;
                }
            }
            currRowNum++;
            _dataRowCount++;
        }

        //if (!execResult.Result)
        //{
        //    SendAppTraceExec(AppTraceLevel.Error, "ExecInstrForEachRowIfThenMgr.Exec", InstrForEachRowIfThenExecEvent.CreateFinishedError(execStart, error));
        //    return false;
        //}

        SendAppTraceExec(AppTraceLevel.Info, "ExecInstrForEachRowIfThenMgr.Exec", InstrForEachRowIfThenExecEvent.CreateFinishedOk(execStart, _dataRowCount, _ifConditionFiredCount));

        return true;
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
    static bool ExecOnDataRow(ExecResult execResult, DateTime execStart, IExcelProcessor excelProcessor, IExcelFile excelFile, IExcelSheet excelSheet, InstrRetBoolBase instrIf, List<InstrBase> listInstrThen, int rowNum)
    {
        // execute If cond on the cells of the datarow
        bool res= ExecIfCondition(execResult, excelProcessor, excelFile, excelSheet, instrIf, rowNum, out bool condResult); 
        if(!res)
        {
            // stop on error, not an warning
            SendAppTraceExec(AppTraceLevel.Error, "ExecInstrForEachRowIfThenMgr.ExecOnDataRow", InstrForEachRowIfThenExecEvent.CreateFinishedError(execStart, execResult.ListError.FirstOrDefault()));
            // error occurs during the If condition instr execution
            return false;
        }

        // the If condition return false, so doesn't execute Then instructions
        if(!condResult)
            return true;

        _ifConditionFiredCount++;
        SendAppTraceExec(AppTraceLevel.Info, "ExecInstrForEachRowIfThenMgr.ExecOnDataRow", InstrForEachRowIfThenExecEvent.CreateFinishedInProgress(execStart, _dataRowCount, _ifConditionFiredCount));

        // execute Then instructions
        foreach (InstrBase instrThen in listInstrThen)
        {
            ExecOnCellInstrThen(execResult, excelProcessor, excelSheet, instrThen, rowNum);
        }

        if (!execResult.Result)
            return false;

        // ok, finish processing all datarow with success
        return true;
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
    static bool ExecIfCondition(ExecResult execResult, IExcelProcessor excelProcessor, IExcelFile excelFile, IExcelSheet excelSheet, InstrRetBoolBase instrIf, int rowNum, out bool condResult)
    {
        condResult = false;

        //--is it: If A.Cell=12 ?
        InstrCompColCellVal instrCompCellVal = instrIf as InstrCompColCellVal;
        if (instrCompCellVal != null)
        {
            return ExecInstrCompColCellVal(execResult, excelProcessor, excelFile, excelSheet, instrCompCellVal, rowNum, out condResult);
        }

        //--is it: If A.Cell=null ?
        InstrCompColCellValIsNull instrCompCellValIsNull = instrIf as InstrCompColCellValIsNull;
        if(instrCompCellValIsNull!=null)
        {
            return ExecInstrCompColCellValIsNull(execResult, excelProcessor, excelFile, excelSheet, instrCompCellValIsNull, rowNum, out condResult);
        }

        //--is it: If A.Cell=blank ?
        // TODO:

        //--is it: If A.Cell=nullblank ?  null or blank
        // TODO:

        //--is it: A.Cell In ["y","yes"] ? 
        InstrCompColCellInStringItems instrCompColCellInStringItems = instrIf as InstrCompColCellInStringItems;
        if(instrCompColCellInStringItems!=null)
        {
            return ExecInstrCompColCellInStringItems(execResult, excelProcessor, excelFile, excelSheet, instrCompColCellInStringItems, rowNum, out condResult);
        }

        //--is it: if A.Cell=12 And B.Cell="Y" And ... ?
        InstrCompListColCellAnd instrCompListColCellAnd = instrIf as InstrCompListColCellAnd;
        if (instrCompListColCellAnd != null)
        {
            return ExecInstrCompListColCellAnd(execResult, excelProcessor, excelFile, excelSheet, instrCompListColCellAnd, rowNum, out condResult);
        }

        //--instr Not allowed as an If condition
        execResult.AddError(new ExecResultError(ErrorCode.IfConditionInstrNotAllowed, instrIf.InstrType.ToString()));
        return false;
    }

    /// <summary>
    /// If instr is a basic comparison, exp: A.Cell=12
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="excelFile"></param>
    /// <param name="excelSheet"></param>
    /// <param name="instrCompColCellVal"></param>
    /// <param name="rowNum"></param>
    /// <param name="condResult"></param>
    /// <returns></returns>
    static bool ExecInstrCompColCellVal(ExecResult execResult, IExcelProcessor excelProcessor, IExcelFile excelFile, IExcelSheet excelSheet, InstrCompColCellVal instrCompColCellVal, int rowNum, out bool condResult)
    {
        condResult = false;

        var cell = excelProcessor.GetCellAt(excelSheet, rowNum, instrCompColCellVal.ColNum);

        // get the cell value type            
        CellRawValueType cellType = excelProcessor.GetCellValueType(excelSheet, cell);

        // does the cell type match the If-Comparison cell.Value type?
        if (!ExcelExtendedUtils.MatchCellTypeAndIfComparison(cellType, instrCompColCellVal.Value))
        {
            // is there an warning already existing? 
            execResult.AddWarning(ErrorCode.IfCondTypeMismatch, excelFile.FileName, excelSheet.Index, instrCompColCellVal.ColNum, cellType);
            return false;
        }

        // execute the If part: comparison condition
        return ExecInstrCompMgr.ExecInstrCompCellVal(execResult, excelProcessor, instrCompColCellVal, cell, out condResult);
    }

    /// <summary>
    /// Is instr if a comparison with null?
    /// exp: A.Cell=null
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="excelFile"></param>
    /// <param name="excelSheet"></param>
    /// <param name="instrCompColCellValIsNull"></param>
    /// <param name="rowNum"></param>
    /// <param name="condResult"></param>
    /// <returns></returns>
    static bool ExecInstrCompColCellValIsNull(ExecResult execResult, IExcelProcessor excelProcessor, IExcelFile excelFile, IExcelSheet excelSheet, InstrCompColCellValIsNull instrCompColCellValIsNull, int rowNum, out bool condResult)
    {
        var cell = excelProcessor.GetCellAt(excelSheet, rowNum, instrCompColCellValIsNull.ColNum);

        // execute the If part: comparison condition
        return ExecInstrCompMgr.ExecInstrCompCellValIsNull(execResult, excelProcessor, instrCompColCellValIsNull, cell, out condResult);
    }

    /// <summary>
    /// If A.Cell In [ "y", "yes", "ok"]
    /// the cell type should be string.
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="excelFile"></param>
    /// <param name="excelSheet"></param>
    /// <param name="instrCompListColCellAnd"></param>
    /// <param name="rowNum"></param>
    /// <param name="condResult"></param>
    /// <returns></returns>
    static bool ExecInstrCompColCellInStringItems(ExecResult execResult, IExcelProcessor excelProcessor, IExcelFile excelFile, IExcelSheet excelSheet, InstrCompColCellInStringItems instr, int rowNum, out bool condResult)
    {
        condResult = false;

        var cell = excelProcessor.GetCellAt(excelSheet, rowNum, instr.ColNum);

        // get the cell value type            
        CellRawValueType cellType = excelProcessor.GetCellValueType(excelSheet, cell);

        // the cell type should be string
        if (cellType != CellRawValueType.String)
        {
            // is there an warning already existing? Code=IfCondTypeMismatch, ExcelFile, fileName, SheetNum, colNum, valType
            execResult.AddWarning(ErrorCode.IfCondTypeMismatch, excelFile.FileName, excelSheet.Index, instr.ColNum, cellType);
            return false;
        }

        // execute the If part: comparison condition
        condResult = ExecInstrCompMgr.ExecInstrCompColCellInStringItems(excelProcessor, instr, cell);
        return true;
    }

    /// <summary>
    /// If A.Cell=12 And B.Cell="Y" And ...
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="excelFile"></param>
    /// <param name="excelSheet"></param>
    /// <param name="instrCompListColCellAnd"></param>
    /// <param name="rowNum"></param>
    /// <param name="condResult"></param>
    /// <returns></returns>
    static bool ExecInstrCompListColCellAnd(ExecResult execResult, IExcelProcessor excelProcessor, IExcelFile excelFile, IExcelSheet excelSheet, InstrCompListColCellAnd instrCompListColCellAnd, int rowNum, out bool condResult)
    {
        bool condResultLocal;
        condResult = true;

        foreach (InstrRetBoolBase instrRetBoolBase in instrCompListColCellAnd.ListInstrComp)
        {
            // stop when the condition result becomes false
            if (!condResult)
                break;

            //--is it: A.Cell=12 ?
            InstrCompColCellVal instrCompCellVal = instrRetBoolBase as InstrCompColCellVal;
            if (instrCompCellVal != null)
            {
                // TODO: gestion des ExecResultWarning!! 
                ExecInstrCompColCellVal(execResult, excelProcessor, excelFile, excelSheet, instrCompCellVal, rowNum, out condResultLocal);
                //if (!execResultLocal.Result)
                    // TODO: gestion error ou warning!
                  //  execResult.AddListError(execResultLocal.ListError);
                condResult &= condResultLocal;
                continue;
            }

            // is it: A.Cell=null ?
            InstrCompColCellValIsNull instrCompCellValIsNull = instrRetBoolBase as InstrCompColCellValIsNull;
            if (instrCompCellValIsNull != null)
            {
                // TODO: gestion des ExecResultWarning!! 
                ExecInstrCompColCellValIsNull(execResult, excelProcessor, excelFile, excelSheet, instrCompCellValIsNull, rowNum, out condResultLocal);
                //if (!execResultLocal.Result)
                    // TODO: gestion error ou warning!
                  //  execResult.AddListError(execResultLocal.ListError);
                condResult &= condResultLocal;
            }

            //--is it: A.Cell In [ "y", "yes", "ok"] ?
            // TODO:
        }

        return true;
    }

    /// <summary>
    /// Execute the then instruction.
    /// InstrSetCellVal, ExecSetCellNull ExecSetCellValueBlank are authorized.
    /// exp: Cell.Val= 12, Cell.Val= null, Cell.Val= blank
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instr"></param>
    /// <param name="sheet"></param>
    /// <param name="colNum"></param>
    /// <param name="rowNum"></param>
    /// <returns></returns>
    static bool ExecOnCellInstrThen(ExecResult execResult, IExcelProcessor excelProcessor, IExcelSheet sheet, InstrBase instr, int rowNum)
    {
        // check allowed Then instructions

        InstrSetCellVal instrSetCellVal = instr as InstrSetCellVal;
        if(instrSetCellVal!=null)
            return  ExecInstrSetCellMgr.ExecSetCellVal(execResult, excelProcessor, instrSetCellVal, sheet, rowNum);

        InstrSetCellNull instrSetCellNull = instr as InstrSetCellNull;
        if(instrSetCellNull!=null)
            return ExecInstrSetCellMgr.ExecSetCellNull(execResult, excelProcessor, instrSetCellNull, sheet, rowNum);

        InstrSetCellValueBlank instrSetCellValueBlank = instr as InstrSetCellValueBlank;
        if (instrSetCellValueBlank != null)
            return ExecInstrSetCellMgr.ExecSetCellValueBlank(execResult, excelProcessor, instrSetCellValueBlank, sheet, rowNum);

        execResult.AddError(new ExecResultError(ErrorCode.ThenConditionInstrNotAllowed, instr.ToString()));
        return false;
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
