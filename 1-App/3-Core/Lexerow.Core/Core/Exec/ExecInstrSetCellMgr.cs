using Lexerow.Core.System.Excel;
using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.Formula.Functions;
using Lexerow.Core.Utils;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.XSSF.Streaming.Values;
using static NPOI.HSSF.Util.HSSFColor;

namespace Lexerow.Core;

/// <summary>
/// Execute the instruction: Set Cell Value.
/// </summary>
public class ExecInstrSetCellMgr
{
    /// <summary>
    /// Execute the instr: Set cell value.
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instr"></param>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    public static ExecResult ExecSetCellVal(IExcelProcessor excelProcessor, InstrSetCellVal instrSetCellVal, IExcelSheet sheet, int rowNum)
    {
        ExecResult execResult = new ExecResult();

        // get the cell
        IExcelCell cell = excelProcessor.GetCellAt(sheet, rowNum, instrSetCellVal.ColNum);

        if(cell!=null) 
            return ExecCellExists(excelProcessor, instrSetCellVal, sheet, rowNum, cell);

        // create a new cell object
        cell = excelProcessor.CreateCell(sheet, rowNum, instrSetCellVal.ColNum);

        return ApplySetCellValAndType(excelProcessor, sheet, cell, instrSetCellVal.Value);
    }

    /// <summary>
    /// Execute the instr: Set cell null.
    /// remove the cell from the sheet. 
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instr"></param>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    public static ExecResult ExecSetCellNull(IExcelProcessor excelProcessor, InstrSetCellNull instrSetCellNull, IExcelSheet sheet, int rowNum)
    {
        ExecResult execResult = new ExecResult();

        // get the cell
        IExcelCell cell = excelProcessor.GetCellAt(sheet, rowNum, instrSetCellNull.ColNum);

        if (cell == null)
            return execResult;

        // create a new cell object
        excelProcessor.DeleteCell(sheet, rowNum, instrSetCellNull.ColNum);
        return execResult;
    }

    public static ExecResult ExecSetCellValueBlank(IExcelProcessor excelProcessor, InstrSetCellValueBlank instrSetCellBlank, IExcelSheet sheet, int rowNum)
    {
        ExecResult execResult = new ExecResult();

        // get the cell
        IExcelCell cell = excelProcessor.GetCellAt(sheet, rowNum, instrSetCellBlank.ColNum);

        if (cell == null)
            return execResult;

        // create a new cell object
        excelProcessor.SetCellValueBlank(cell);
        return execResult;
    }

    public static ExecResult ExecCellExists(IExcelProcessor excelProcessor, InstrSetCellVal instrSetCellVal, IExcelSheet sheet, int rowNum, IExcelCell cell)
    {
        // get the cell value type
        CellRawValueType cellType = excelProcessor.GetCellValueType(sheet, cell);

        // does the setCellVal and cell value match?
        bool res = ExcelExtendedUtils.MatchCellTypeAndIfComparison(cellType, instrSetCellVal.Value);

        // yes
        if (res)
            return ApplySetCellVal(excelProcessor, sheet, cell, instrSetCellVal.Value);

        // type mismatch:problem, the cell exists but the value type to set is different
        return ApplySetCellValAndType(excelProcessor, sheet, cell, instrSetCellVal.Value);
    }

    static ExecResult ApplySetCellVal(IExcelProcessor excelProcessor, IExcelSheet sheet, IExcelCell cell, ValueBase value)
    {
        ExecResult execResult = new ExecResult();

        if (value.ValueType == System.ValueType.Int)
        {
            // set the new value to the cell
            excelProcessor.SetCellValue(cell, (value as ValueInt).Val);
            return execResult;
        }
        if (value.ValueType == System.ValueType.Double)
        {
            // set the new value to the cell
            excelProcessor.SetCellValue(cell, (value as ValueDouble).Val);
            return execResult;
        }

        if (value.ValueType == System.ValueType.String)
        {
            // set the new value to the cell
            excelProcessor.SetCellValue(cell, (value as ValueString).Val);
            return execResult;
        }

        if (value.ValueType == System.ValueType.DateOnly)
        {
            // set the new value to the cell
            excelProcessor.SetCellValue(cell, (value as ValueDateOnly).ToDouble());
            return execResult;
        }

        if (value.ValueType == System.ValueType.TimeOnly)
        {
            // set the new value to the cell
            excelProcessor.SetCellValue(cell, (value as ValueTimeOnly).ToDouble());
            return execResult;
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
            return execResult;
        }

        // type not managed
        execResult.AddError(new ExecResultError(ErrorCode.ExcelUnableOpenFile, value.ValueType.ToString()));
        return execResult;
    }

    static ExecResult ApplySetCellValAndType(IExcelProcessor excelProcessor, IExcelSheet sheet, IExcelCell cell, ValueBase value)
    {
        ExecResult execResult = new ExecResult();

        if (value.ValueType == System.ValueType.String)
        {
            excelProcessor.SetCellValueString(cell, (value as ValueString).Val);
            return execResult;
        }

        if (value.ValueType == System.ValueType.Int)
        {
            excelProcessor.SetCellValueInt(cell, (value as ValueInt).Val);
            return execResult;
        }

        if (value.ValueType == System.ValueType.Double)
        {
            excelProcessor.SetCellValueDouble(cell, (value as ValueDouble).Val);
            return execResult;
        }

        if (value.ValueType == System.ValueType.DateOnly)
        {
            excelProcessor.SetCellValueDateOnly(cell, value as ValueDateOnly);
            return execResult;
        }

        if (value.ValueType == System.ValueType.DateTime)
        {
            excelProcessor.SetCellValueDateTime(cell, value as ValueDateTime);
            return execResult;
        }

        if (value.ValueType == System.ValueType.TimeOnly)
        {
            excelProcessor.SetCellValueTimeOnly(cell, value as ValueTimeOnly);
            return execResult;
        }

        // type not managed
        execResult.AddError(new ExecResultError(ErrorCode.ExcelCellTypeNotManaged, value.ValueType.ToString()));
        return execResult;
    }

}
