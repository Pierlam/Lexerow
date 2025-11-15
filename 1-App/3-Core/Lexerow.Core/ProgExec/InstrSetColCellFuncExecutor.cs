using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.Utils;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgExec;
public class InstrSetColCellFuncExecutor
{
    IActivityLogger _logger;

    IExcelProcessor _excelProcessor;

    public InstrSetColCellFuncExecutor(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
    }

    /// <summary>
    /// A.Cell= 12
    /// TODO: see ExecInstrSetCellMgr
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="instrColCellFunc"></param>
    /// <param name="instrRight"></param>
    /// <returns></returns>
    public bool ExecSetCellValue(ExecResult execResult, IExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFunc, InstrConstValue instrConstValue)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrSetColCellFuncRunner.Run", string.Empty);

        // get the cell
        IExcelCell cell = _excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFunc.ColNum-1);

        if (cell != null)
            return ExecCellExists(execResult, _excelProcessor, excelSheet, rowNum, instrConstValue, cell);

        // create a new cell object
        cell = _excelProcessor.CreateCell(excelSheet, rowNum, instrColCellFunc.ColNum-1);

        return ApplySetCellValAndType(execResult, _excelProcessor, excelSheet, cell, instrConstValue.ValueBase);
    }

    /// <summary>
    /// Execute the instr: Set cell null.
    /// remove the cell from the sheet. 
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instr"></param>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    public bool ExecSetCellNull(ExecResult execResult, IExcelSheet sheet, int rowNum, InstrColCellFunc instrColCellFunc)
    {
        // get the cell
        IExcelCell cell = _excelProcessor.GetCellAt(sheet, rowNum, instrColCellFunc.ColNum-1);

        if (cell == null)
            return true;

        // create a new cell object
        _excelProcessor.DeleteCell(sheet, rowNum, instrColCellFunc.ColNum-1);
        return true;
    }

    public bool ExecSetCellBlank(ExecResult execResult, IExcelSheet sheet, int rowNum, InstrColCellFunc instrColCellFunc)
    {
        // get the cell
        IExcelCell cell = _excelProcessor.GetCellAt(sheet, rowNum, instrColCellFunc.ColNum - 1);

        if (cell == null)
            return true;

        // create a new cell object
        _excelProcessor.SetCellValueBlank(cell);
        return true;
    }

    bool ExecCellExists(ExecResult execResult, IExcelProcessor excelProcessor, IExcelSheet sheet, int rowNum, InstrConstValue instrSetCellVal, IExcelCell cell)
    {
        // get the cell value type
        CellRawValueType cellType = excelProcessor.GetCellValueType(sheet, cell);

        // does the setCellVal and cell value match?
        bool res = ExcelExtendedUtils.MatchCellTypeAndIfComparison(cellType, instrSetCellVal.ValueBase);

        // yes
        if (res)
            return ApplySetCellVal(execResult, excelProcessor, sheet, cell, instrSetCellVal.ValueBase);

        // type mismatch:problem, the cell exists but the value type to set is different
        return ApplySetCellValAndType(execResult, excelProcessor, sheet, cell, instrSetCellVal.ValueBase);
    }

    bool ApplySetCellVal(ExecResult execResult, IExcelProcessor excelProcessor, IExcelSheet sheet, IExcelCell cell, ValueBase value)
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
        execResult.AddError(new ExecResultError(ErrorCode.ExcelUnableOpenFile, value.ValueType.ToString()));
        return false;
    }

    bool ApplySetCellValAndType(ExecResult execResult, IExcelProcessor excelProcessor, IExcelSheet sheet, IExcelCell cell, ValueBase value)
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
        execResult.AddError(ErrorCode.ExcelCellTypeNotManaged, value.ValueType.ToString());
        return false;
    }

}
