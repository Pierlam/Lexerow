using DocumentFormat.OpenXml.Spreadsheet;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.Utils;
using Microsoft.VisualBasic;
using OpenExcelSdk;
using OpenExcelSdk.System;

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
    public bool ExecSetCellNull(Result result, ExcelSheet sheet, int rowNum, InstrColCellFunc instrColCellFunc)
    {
        // get the cell
        //IExcelCell cell = _excelProcessor.GetCellAt(sheet, rowNum, instrColCellFunc.ColNum - 1);
        _excelProcessor.GetCellAt(sheet, instrColCellFunc.ColNum, rowNum, out ExcelCell cell, out ExcelError error);

        if (cell == null)
            return true;

        // create a new cell object
        _excelProcessor.RemoveCell(sheet, instrColCellFunc.ColNum, rowNum, out error);
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
        //IExcelCell cell = _excelProcessor.GetCellAt(sheet, rowNum, instrColCellFunc.ColNum - 1);
        _excelProcessor.GetCellAt(sheet, instrColCellFunc.ColNum, rowNum, out ExcelCell cell, out ExcelError error);

        if (cell == null)
            return true;

        // create a new cell object
        //_excelProcessor.SetCellValueBlank(cell);
        throw new Exception("TODO: SetCellEmpty()");
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
    public bool ExecSetCellValue(Result result, ExcelSheet sheet, int rowNum, InstrColCellFunc instrColCellFunc, InstrValue instrValue)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrSetColCellFuncExecutor.ExecSetCellValue", string.Empty);

        // get the cell
        //IExcelCell cell = _excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFunc.ColNum - 1);
        _excelProcessor.GetCellAt(sheet, instrColCellFunc.ColNum, rowNum, out ExcelCell cell, out ExcelError error);
        if (cell == null)
            _excelProcessor.CreateCell(sheet, instrColCellFunc.ColNum, rowNum, out cell, out error);

        return ApplySetCellVal(result, _excelProcessor, sheet, cell, instrValue.ValueBase);
        //_excelProcessor.SetCellValue(sheet, cell, )

        //if (cell == null)
        //    // create a new cell object
        //    cell = _excelProcessor.CreateCell(excelSheet, rowNum, instrColCellFunc.ColNum - 1);

        //// get the cell value type
        //CellRawValueType cellType = _excelProcessor.GetCellValueType(excelSheet, cell);

        //// does the setCellVal and cell value match?
        //bool res = ExcelExtendedUtils.MatchCellTypeAndIfComparison(cellType, instrValue.ValueBase);

        //// yes
        //if (res)
        //    return ApplySetCellVal(result, _excelProcessor, excelSheet, cell, instrValue.ValueBase);

        //return ApplySetCellValAndType(result, _excelProcessor, excelSheet, cell, instrValue.ValueBase);
        return false;
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
    public bool ExecSetCellValue_REWORK(Result result, ExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFunc, InstrValue instrValue)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrSetColCellFuncExecutor.ExecSetCellValue", string.Empty);

        // if the row doesn't exists!!
        // TODO:

        // get the cell
        //IExcelCell cell = _excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFunc.ColNum - 1);

        //if(cell == null)
        //    // create a new cell, type is blank, a default style is created and set
        //    cell = _excelProcessor.CreateCell(excelSheet, rowNum, instrColCellFunc.ColNum - 1);


        //// get the cell value type
        //CellRawValueType cellType = _excelProcessor.GetCellValueType(excelSheet, cell);

        //// the value to set is a string, exp: A.Cell="hello"
        //if (instrValue.ValueBase.ValueType == System.ValueType.String)
        //    return SetCellValueString(result, _excelProcessor, excelSheet, rowNum, instrColCellFunc.ColNum - 1, cell, cellType, (instrValue.ValueBase as ValueString).Val);

        // TODO: Number: Int/Double

        // TODO: Date

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
    private bool ApplySetCellVal(Result result, ExcelProcessor excelProcessor, ExcelSheet sheet, ExcelCell cell, ValueBase value)
    {
        ExcelError error;

        if (value.ValueType == System.ValueType.Int)
        {
            // set the new value to the cell
            if(!excelProcessor.SetCellValue(sheet, cell, (value as ValueInt).Val, out error))
                return result.AddError(ErrorUtils.Convert(error));
            return true;
        }

        if (value.ValueType == System.ValueType.Double)
        {
            // set the new value to the cell
            if(!excelProcessor.SetCellValue(sheet, cell, (value as ValueDouble).Val, out error))
                return result.AddError(ErrorUtils.Convert(error));
            return true;
        }

        if (value.ValueType == System.ValueType.String)
        {
            // set the new value to the cell
            if(!excelProcessor.SetCellValue(sheet, cell, (value as ValueString).Val, out error))
                return result.AddError(ErrorUtils.Convert(error));
            return true;
        }

        if (value.ValueType == System.ValueType.DateOnly)
        {
            string sysVarDateFormat = _progExecVarMgr.GetProgExecSysVarAsString(CoreInstr.SysVarDateFormatName);

            // set the new value to the cell
            if (!excelProcessor.SetCellValue(sheet, cell, (value as ValueDateOnly).Val, sysVarDateFormat, out error))
                return result.AddError(ErrorUtils.Convert(error));

            return true;
        }

        if (value.ValueType == System.ValueType.DateTime)
        {
            string sysVarDateTimeFormat = _progExecVarMgr.GetProgExecSysVarAsString(CoreInstr.SysVarDateTimeFormatName);

            // set the new value to the cell
            if (!excelProcessor.SetCellValue(sheet, cell, (value as ValueDateTime).Val, sysVarDateTimeFormat, out error))
                return result.AddError(ErrorUtils.Convert(error));

            return true;
        }

        if (value.ValueType == System.ValueType.TimeOnly)
        {
            string sysVarTimeFormat = _progExecVarMgr.GetProgExecSysVarAsString(CoreInstr.SysVarDateTimeFormatName);
            // set the new value to the cell
            if (!excelProcessor.SetCellValue(sheet, cell, (value as ValueTimeOnly).Val, sysVarTimeFormat, out error))
                return result.AddError(ErrorUtils.Convert(error));

            return true;
        }

        // type not managed
        result.AddError(new ResultError(ErrorCode.ExcelUnableOpenFile, value.ValueType.ToString()));
        return false;
    }
}