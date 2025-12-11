using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.Utils;
using Microsoft.VisualBasic;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Execute instr set value to a ColCellFunc.
/// exp: A.Cell=10, A.Cell=null, A.Cell=a,...
/// </summary>
public class InstrSetColCellFuncExecutor
{
    private IActivityLogger _logger;

    private IExcelProcessor _excelProcessor;

    private ProgExecVarMgr _progExecVarMgr;

    public InstrSetColCellFuncExecutor(IActivityLogger activityLogger, IExcelProcessor excelProcessor, ProgExecVarMgr progExecVarMgr)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
        _progExecVarMgr= progExecVarMgr;
    }

    /// <summary>
    /// Execute the instr: Set cell null.
    /// remove the cell from the sheet.
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instr"></param>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    public bool ExecSetCellNull(Result result, IExcelSheet sheet, int rowNum, InstrColCellFunc instrColCellFunc)
    {
        // get the cell
        IExcelCell cell = _excelProcessor.GetCellAt(sheet, rowNum, instrColCellFunc.ColNum - 1);

        if (cell == null)
            return true;

        // create a new cell object
        _excelProcessor.DeleteCell(sheet, rowNum, instrColCellFunc.ColNum - 1);
        return true;
    }

    /// <summary>
    /// Execute Set Cell=blank
    /// Keep the formatting of the cell.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="sheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="instrColCellFunc"></param>
    /// <returns></returns>
    public bool ExecSetCellBlank(Result result, IExcelSheet sheet, int rowNum, InstrColCellFunc instrColCellFunc)
    {
        // get the cell
        IExcelCell cell = _excelProcessor.GetCellAt(sheet, rowNum, instrColCellFunc.ColNum - 1);

        if (cell == null)
            return true;

        // create a new cell object
        _excelProcessor.SetCellValueBlank(cell);
        return true;
    }

    /// <summary>
    /// Set a value to a cell.
    /// exp:A.Cell=12
    /// </summary>
    /// <param name="result"></param>
    /// <param name="instrColCellFunc"></param>
    /// <param name="instrRight"></param>
    /// <returns></returns>
    public bool ExecSetCellValue(Result result, IExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFunc, InstrValue instrValue)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrSetColCellFuncExecutor.ExecSetCellValue", string.Empty);

        // get the cell
        IExcelCell cell = _excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFunc.ColNum - 1);

        if (cell == null)
            // create a new cell object
            cell = _excelProcessor.CreateCell(excelSheet, rowNum, instrColCellFunc.ColNum - 1);

        // get the cell value type
        CellRawValueType cellType = _excelProcessor.GetCellValueType(excelSheet, cell);

        // does the setCellVal and cell value match?
        bool res = ExcelExtendedUtils.MatchCellTypeAndIfComparison(cellType, instrValue.ValueBase);

        // yes
        if (res)
            return ApplySetCellVal(result, _excelProcessor, excelSheet, cell, instrValue.ValueBase);

        return ApplySetCellValAndType(result, _excelProcessor, excelSheet, cell, instrValue.ValueBase);
    }

    /// <summary>
    /// Under construction!
    /// next version to use.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="instrColCellFunc"></param>
    /// <param name="instrValue"></param>
    /// <returns></returns>
    public bool ExecSetCellValue_REWORK(Result result, IExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFunc, InstrValue instrValue)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrSetColCellFuncExecutor.ExecSetCellValue", string.Empty);

        // if the row doesn't exists!!
        // TODO:

        // get the cell
        IExcelCell cell = _excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFunc.ColNum - 1);

        if(cell == null)
            // create a new cell, type is blank, a default style is created and set
            cell = _excelProcessor.CreateCell(excelSheet, rowNum, instrColCellFunc.ColNum - 1);


        // get the cell value type
        CellRawValueType cellType = _excelProcessor.GetCellValueType(excelSheet, cell);

        // the value to set is a string, exp: A.Cell="hello"
        if (instrValue.ValueBase.ValueType == System.ValueType.String)
            return SetCellValueString(result, _excelProcessor, excelSheet, rowNum, instrColCellFunc.ColNum - 1, cell, cellType, (instrValue.ValueBase as ValueString).Val);

        // TODO: Number: Int/Double

        // TODO: Date

        return true;
    }

    /// <summary>
    /// Set a string value to a cell.
    /// exp: A.Cell="hello"
    /// </summary>
    /// <param name="result"></param>
    /// <param name="excelProcessor"></param>
    /// <param name="sheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="colNum"></param>
    /// <param name="cell"></param>
    /// <param name="cellType"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    private bool SetCellValueString(Result result, IExcelProcessor excelProcessor, IExcelSheet sheet, int rowNum, int colNum, IExcelCell cell, CellRawValueType cellType, string value)
    {
        // cell type and value type are identical
        if(cellType == CellRawValueType.String)
            return excelProcessor.SetCellValue(cell, value);

        // type are differents, clone the style of the cell
        // TODO:
        return true;
    }

    /// <summary>
    /// Set a value into a cell that exists.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="excelProcessor"></param>
    /// <param name="sheet"></param>
    /// <param name="cell"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    private bool ApplySetCellVal(Result result, IExcelProcessor excelProcessor, IExcelSheet sheet, IExcelCell cell, ValueBase value)
    {
        if (value.ValueType == System.ValueType.Int)
        {
            // set the new value to the cell
            excelProcessor.SetCellValue(cell, (value as ValueInt).Val);
            return true;
        }

        if (value.ValueType == System.ValueType.Double)
        {
            // set the new value to the cell
            excelProcessor.SetCellValue(cell, (value as ValueDouble).Val);
            return true;
        }

        if (value.ValueType == System.ValueType.String)
        {
            // set the new value to the cell
            excelProcessor.SetCellValue(cell, (value as ValueString).Val);
            return true;
        }

        if (value.ValueType == System.ValueType.DateOnly)
        {
            // set the new value to the cell
            excelProcessor.SetCellValue(cell, (value as ValueDateOnly).ToDouble());
            return true;
        }

        if (value.ValueType == System.ValueType.TimeOnly)
        {
            // set the new value to the cell
            excelProcessor.SetCellValue(cell, (value as ValueTimeOnly).ToDouble());
            return true;
        }

        if (value.ValueType == System.ValueType.DateTime)
        {
            double dbVal = (value as ValueDateTime).ToDouble();

            CellRawValueType cellType = excelProcessor.GetCellValueType(sheet, cell);
            if (cellType == CellRawValueType.DateOnly)
                // the cell is a dataOnly but the value to set is a DateTime -> remove the hour before set.
                dbVal = Math.Truncate(dbVal);

            // set the new value to the cell
            excelProcessor.SetCellValue(cell, dbVal);
            return true;
        }

        // type not managed
        result.AddError(new ResultError(ErrorCode.ExcelUnableOpenFile, value.ValueType.ToString()));
        return false;
    }

    private bool ApplySetCellValAndType(Result result, IExcelProcessor excelProcessor, IExcelSheet sheet, IExcelCell cell, ValueBase value)
    {
        if (value.ValueType == System.ValueType.String)
        {
            excelProcessor.SetCellValue(cell, (value as ValueString).Val);
            return true;
        }

        if (value.ValueType == System.ValueType.Int)
        {
            excelProcessor.SetCellValue(cell, (value as ValueInt).Val);
            return true;
        }

        if (value.ValueType == System.ValueType.Double)
        {
            excelProcessor.SetCellValue(cell, (value as ValueDouble).Val);
            return true;
        }

        if (value.ValueType == System.ValueType.DateOnly)
        {
            string sysVarDateFormat = _progExecVarMgr.GetProgExecSysVarAsString(CoreInstr.SysVarDateFormatName);
            bool sysVarForceDateFormat = _progExecVarMgr.GetProgExecSysVarAsBool(CoreInstr.SysVarForceDateFormatName);

            excelProcessor.SetCellValueDateOnly(cell, value as ValueDateOnly);
            return true;
        }

        if (value.ValueType == System.ValueType.DateTime)
        {
            excelProcessor.SetCellValueDateTime(cell, value as ValueDateTime);
            return true;
        }

        if (value.ValueType == System.ValueType.TimeOnly)
        {
            excelProcessor.SetCellValueTimeOnly(cell, value as ValueTimeOnly);
            return true;
        }

        // type not managed
        result.AddError(ErrorCode.ExcelCellTypeNotManaged, value.ValueType.ToString());
        return false;
    }
}