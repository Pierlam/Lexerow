using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.Utils;

namespace Lexerow.Core.InstrProgExec;

public class InstrSetColCellFuncExecutor
{
    private IActivityLogger _logger;

    private IExcelProcessor _excelProcessor;

    public InstrSetColCellFuncExecutor(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
    }

    /// <summary>
    /// A.Cell= 12
    /// TODO: see ExecInstrSetCellMgr
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

        if (cell != null)
            return ExecCellExists(result, _excelProcessor, excelSheet, rowNum, instrValue, cell);

        // create a new cell object
        cell = _excelProcessor.CreateCell(excelSheet, rowNum, instrColCellFunc.ColNum - 1);

        return ApplySetCellValAndType(result, _excelProcessor, excelSheet, cell, instrValue.ValueBase);
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

    private bool ExecCellExists(Result result, IExcelProcessor excelProcessor, IExcelSheet sheet, int rowNum, InstrValue instrSetCellVal, IExcelCell cell)
    {
        // get the cell value type
        CellRawValueType cellType = excelProcessor.GetCellValueType(sheet, cell);

        // does the setCellVal and cell value match?
        bool res = ExcelExtendedUtils.MatchCellTypeAndIfComparison(cellType, instrSetCellVal.ValueBase);

        // yes
        if (res)
            return ApplySetCellVal(result, excelProcessor, sheet, cell, instrSetCellVal.ValueBase);

        // type mismatch:problem, the cell exists but the value type to set is different
        return ApplySetCellValAndType(result, excelProcessor, sheet, cell, instrSetCellVal.ValueBase);
    }

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
            excelProcessor.SetCellValueString(cell, (value as ValueString).Val);
            return true;
        }

        if (value.ValueType == System.ValueType.Int)
        {
            excelProcessor.SetCellValueInt(cell, (value as ValueInt).Val);
            return true;
        }

        if (value.ValueType == System.ValueType.Double)
        {
            excelProcessor.SetCellValueDouble(cell, (value as ValueDouble).Val);
            return true;
        }

        if (value.ValueType == System.ValueType.DateOnly)
        {
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