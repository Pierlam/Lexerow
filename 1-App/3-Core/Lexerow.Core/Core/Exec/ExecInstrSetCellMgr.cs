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

namespace Lexerow.Core.Core.Exec;

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
    public static bool ExecSetCellVal(ExecResult execResult, IExcelProcessor excelProcessor, InstrSetCellVal instrSetCellVal, IExcelSheet sheet, int rowNum)
    {
        // get the cell
        IExcelCell cell = excelProcessor.GetCellAt(sheet, rowNum, instrSetCellVal.ColNum);

        if(cell!=null) 
            return ExecCellExists(execResult, excelProcessor, instrSetCellVal, sheet, rowNum, cell);

        // create a new cell object
        cell = excelProcessor.CreateCell(sheet, rowNum, instrSetCellVal.ColNum);

        return ApplySetCellValAndType(execResult, excelProcessor, sheet, cell, instrSetCellVal.Value);
    }

    /// <summary>
    /// Execute the instr: Set cell null.
    /// remove the cell from the sheet. 
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instr"></param>
    /// <param name="excelFile"></param>
    /// <returns></returns>
    public static bool ExecSetCellNull(ExecResult execResult, IExcelProcessor excelProcessor, InstrSetCellNull instrSetCellNull, IExcelSheet sheet, int rowNum)
    {
        // get the cell
        IExcelCell cell = excelProcessor.GetCellAt(sheet, rowNum, instrSetCellNull.ColNum);

        if (cell == null)
            return true;

        // create a new cell object
        excelProcessor.DeleteCell(sheet, rowNum, instrSetCellNull.ColNum);
        return true;
    }

    public static bool ExecSetCellValueBlank(ExecResult execResult, IExcelProcessor excelProcessor, InstrSetCellValueBlank instrSetCellBlank, IExcelSheet sheet, int rowNum)
    {
        // get the cell
        IExcelCell cell = excelProcessor.GetCellAt(sheet, rowNum, instrSetCellBlank.ColNum);

        if (cell == null)
            return true;

        // create a new cell object
        excelProcessor.SetCellValueBlank(cell);
        return true;
    }

    public static bool ExecCellExists(ExecResult execResult, IExcelProcessor excelProcessor, InstrSetCellVal instrSetCellVal, IExcelSheet sheet, int rowNum, IExcelCell cell)
    {
        // get the cell value type
        CellRawValueType cellType = excelProcessor.GetCellValueType(sheet, cell);

        // does the setCellVal and cell value match?
        bool res = ExcelExtendedUtils.MatchCellTypeAndIfComparison(cellType, instrSetCellVal.Value);

        // yes
        if (res)
            return ApplySetCellVal(execResult, excelProcessor, sheet, cell, instrSetCellVal.Value);

        // type mismatch:problem, the cell exists but the value type to set is different
        return ApplySetCellValAndType(execResult, excelProcessor, sheet, cell, instrSetCellVal.Value);
    }

    static bool ApplySetCellVal(ExecResult execResult, IExcelProcessor excelProcessor, IExcelSheet sheet, IExcelCell cell, ValueBase value)
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

    static bool ApplySetCellValAndType(ExecResult execResult, IExcelProcessor excelProcessor, IExcelSheet sheet, IExcelCell cell, ValueBase value)
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
        execResult.AddError(new ExecResultError(ErrorCode.ExcelCellTypeNotManaged, value.ValueType.ToString()));
        return false;
    }

}
