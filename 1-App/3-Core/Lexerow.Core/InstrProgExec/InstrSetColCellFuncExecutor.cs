using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using OpenExcelSdk;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Execute instr set value to a ColCellFunc.
/// exp: A.Cell=10, A.Cell=null, A.Cell=a,...
/// </summary>
public class InstrSetColCellFuncExecutor
{
    private IActivityLogger _logger;

    private ExcelProcessor _excelProcessor;

    private ProgExecVarMgr _progExecVarMgr;

    public InstrSetColCellFuncExecutor(IActivityLogger activityLogger, ExcelProcessor excelProcessor, ProgExecVarMgr progExecVarMgr)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
        _progExecVarMgr = progExecVarMgr;
    }

    /// <summary>
    /// Execute the instr: Set cell null.
    /// remove the cell from the sheet.
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instr"></param>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    public bool ExecSetCellNull(Result result, ExcelSheet sheet, int rowNum, InstrColCellFunc instrColCellFunc)
    {
        // get the cell
        ExcelCell cell = _excelProcessor.GetCellAt(sheet, instrColCellFunc.ColNum, rowNum);

        if (cell == null)
            return true;

        // create a new cell object
        _excelProcessor.RemoveCell(sheet, instrColCellFunc.ColNum, rowNum);
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
    public bool ExecSetCellBlank(Result result, ExcelSheet sheet, int rowNum, InstrColCellFunc instrColCellFunc)
    {
        // get the cell
        ExcelCell cell = _excelProcessor.GetCellAt(sheet, instrColCellFunc.ColNum, rowNum);

        if (cell == null)
            return true;

        // clear the cell value
        return _excelProcessor.SetCellValue(sheet, cell, string.Empty);
    }

    /// <summary>
    /// Set a value to a cell.
    /// exp:A.Cell=12
    /// </summary>
    /// <param name="result"></param>
    /// <param name="instrColCellFunc"></param>
    /// <param name="instrRight"></param>
    /// <returns></returns>
    public bool ExecSetCellValue(Result result, ExcelSheet sheet, int rowNum, InstrColCellFunc instrColCellFunc, InstrValue instrValue)
    {
        _logger.LogExecStart(ActivityLogLevel.Debug, "InstrSetColCellFuncExecutor.ExecSetCellValue", string.Empty);

        // get the cell
        ExcelCell cell = _excelProcessor.GetCellAt(sheet, instrColCellFunc.ColNum, rowNum);
        if (cell == null)
            cell = _excelProcessor.CreateCell(sheet, instrColCellFunc.ColNum, rowNum);

        return ApplySetCellVal(result, _excelProcessor, sheet, cell, instrValue.ValueBase);
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
    private bool ApplySetCellVal(Result result, ExcelProcessor excelProcessor, ExcelSheet sheet, ExcelCell cell, ValueBase value)
    {
        try
        {
            if (value.ValueType == System.ValueType.Int)
                // set the new value to the cell
                return excelProcessor.SetCellValue(sheet, cell, (value as ValueInt).Val);

            if (value.ValueType == System.ValueType.Double)
                // set the new value to the cell
                return excelProcessor.SetCellValue(sheet, cell, (value as ValueDouble).Val);

            if (value.ValueType == System.ValueType.String)
            {
                // set the new value to the cell
                return excelProcessor.SetCellValue(sheet, cell, (value as ValueString).Val);
            }

            if (value.ValueType == System.ValueType.DateOnly)
            {
                string sysVarDateFormat = _progExecVarMgr.GetProgExecSysVarAsString(CoreInstr.SysVarDateFormatName);

                // set the new value to the cell
                return excelProcessor.SetCellValue(sheet, cell, (value as ValueDateOnly).Val, sysVarDateFormat);
            }

            if (value.ValueType == System.ValueType.DateTime)
            {
                string sysVarDateTimeFormat = _progExecVarMgr.GetProgExecSysVarAsString(CoreInstr.SysVarDateTimeFormatName);

                // set the new value to the cell
                return excelProcessor.SetCellValue(sheet, cell, (value as ValueDateTime).Val, sysVarDateTimeFormat);
            }

            if (value.ValueType == System.ValueType.TimeOnly)
            {
                string sysVarTimeFormat = _progExecVarMgr.GetProgExecSysVarAsString(CoreInstr.SysVarDateTimeFormatName);
                // set the new value to the cell
                return excelProcessor.SetCellValue(sheet, cell, (value as ValueTimeOnly).Val, sysVarTimeFormat);
            }
            // type not managed
            result.AddError(ErrorCode.ExcelUnableSetCellValue, value.ValueType.ToString());
            return false;
        }
        catch (Exception ex)
        {
            result.AddError(ErrorCode.ExcelUnableSetCellValue, ex, value.ValueType.ToString());
            return false;
        }
    }
}