using Lexerow.Core.System;
using Lexerow.Core.System.Excel;
using Lexerow.Core.Utils;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static NPOI.HSSF.Util.HSSFColor;

namespace Lexerow.Core;
public class ExecInstrCompMgr
{
    /// <summary>
    /// Execute the comparison condition.
    /// Type of cell value and instr comparison should match.
    /// 
    /// If -condition- Then
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="SheetNum"></param>
    /// <param name="listCols"></param>
    /// <param name="instrComp"></param>
    /// <param name="compResult"></param>
    /// <returns></returns>
    public static ExecResult ExecInstrCompCellVal(IExcelProcessor excelProcessor, InstrCompCellVal instrComp, IExcelCell excelCell, out bool compResult)
    {
        ExecResult execResult = new ExecResult();

        if (instrComp.Value.ValueType == System.ValueType.Int)
        {
            double dblVal = excelCell.GetRawValueNumeric();
            compResult= ExecCompNumeric(instrComp, dblVal, ((ValueInt)instrComp.Value).Val);
            return execResult;  
        }

        if (instrComp.Value.ValueType == System.ValueType.Double)
        {
            double dblVal = excelCell.GetRawValueNumeric();
            compResult = ExecCompNumeric(instrComp, dblVal, ((ValueDouble)instrComp.Value).Val);
            return execResult;
        }

        if (instrComp.Value.ValueType== System.ValueType.String)
        {
            string stringVal= excelCell.GetRawValueString();
            // only Equal or NotEqual is possible
            compResult = ExecCompString(instrComp, stringVal, ((ValueString)instrComp.Value).Val);
            return execResult;
        }

        if (instrComp.Value.ValueType == System.ValueType.DateOnly)
        {
            double doubleVal = excelCell.GetRawValueNumeric();
            DateTime dtVal = DateTime.FromOADate(doubleVal);
            compResult = ExecCompDateTime(instrComp, dtVal, ((ValueDateOnly)instrComp.Value).ToDateTime());
            return execResult;
        }

        if (instrComp.Value.ValueType == System.ValueType.TimeOnly)
        {
            // should be a value less than 0, exp: 0,354746
            double doubleVal = excelCell.GetRawValueNumeric();
            TimeOnly timeOnly= DateTimeUtils.ToTimeOnly(doubleVal);
            compResult = ExecCompTimeOnly(instrComp, timeOnly, ((ValueTimeOnly)instrComp.Value).Val);
            return execResult;
        }

        if (instrComp.Value.ValueType == System.ValueType.DateTime)
        {
            double doubleVal = excelCell.GetRawValueNumeric();
            DateTime dtVal = DateTimeUtils.ToDateTime(doubleVal);
            compResult = ExecCompDateTime(instrComp, dtVal, ((ValueDateTime)instrComp.Value).Val);
            return execResult;
        }

        compResult = false;
        return execResult;
    }

    /// <summary>
    /// Execute the comparison condition.
    /// Only 2 cases: equal or different, exp:
    /// If A.Cell = Null Then
    /// If A.Cell <> Null Then
    /// 
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instrComp"></param>
    /// <param name="excelCell"></param>
    /// <param name="compResult"></param>
    /// <returns></returns>
    public static ExecResult ExecInstrCompCellValIsNull(IExcelProcessor excelProcessor, InstrCompCellValIsNull instrComp, IExcelCell excelCell, out bool compResult)
    {
        ExecResult execResult = new ExecResult();

        if (excelCell == null || excelCell.GetRawValueString().Length==0)
        {
            // if A.Cell = null -> yes
            if (instrComp.Operator == InstrCompCellValOperator.Equal)
            {
                compResult = true;
                return execResult;
            }
            // if A.Cell = null -> no
            if (instrComp.Operator == InstrCompCellValOperator.NotEqual)
            {
                compResult = false;
                return execResult;
            }
            // XX ERROR
            compResult = false;
            return execResult;
        }

        // the cell hasa value
        // if A.Cell <> null -> yes
        if (instrComp.Operator == InstrCompCellValOperator.NotEqual)
        {
            compResult = true;
            return execResult;
        }

        // if A.Cell = null -> no
        if (instrComp.Operator == InstrCompCellValOperator.Equal)
        {
            compResult = false;
            return execResult;
        }

        // XX ERROR
        compResult = false;
            return execResult;

        }

        /// <summary>
        /// CellValue operator value
        /// exp >12
        /// </summary>
        /// <param name="instrComp"></param>
        /// <param name="cellValue"></param>
        /// <param name="valueComp"></param>
        /// <returns></returns>
        static bool ExecCompNumeric(InstrCompCellVal instrComp, double cellValue, double valueComp)
    {
        if(instrComp.Operator== InstrCompCellValOperator.Equal) 
            return cellValue == valueComp;

        if (instrComp.Operator == InstrCompCellValOperator.NotEqual)
            return cellValue != valueComp;

        if (instrComp.Operator == InstrCompCellValOperator.GreaterThan)
            return cellValue > valueComp;
        if (instrComp.Operator == InstrCompCellValOperator.GreaterOrEqualThan)
            return cellValue >= valueComp;

        if (instrComp.Operator == InstrCompCellValOperator.LesserThan)
            return cellValue < valueComp;
        if (instrComp.Operator == InstrCompCellValOperator.LesserOrEqualThan)
            return cellValue <= valueComp;
        return false;
    }

    static bool ExecCompDateTime(InstrCompCellVal instrComp, DateTime cellValue, DateTime valueComp)
    {
        if (instrComp.Operator == InstrCompCellValOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == InstrCompCellValOperator.NotEqual)
            return cellValue != valueComp;

        if (instrComp.Operator == InstrCompCellValOperator.GreaterThan)
            return cellValue > valueComp;
        if (instrComp.Operator == InstrCompCellValOperator.GreaterOrEqualThan)
            return cellValue >= valueComp;

        if (instrComp.Operator == InstrCompCellValOperator.LesserThan)
            return cellValue < valueComp;
        if (instrComp.Operator == InstrCompCellValOperator.LesserOrEqualThan)
            return cellValue <= valueComp;
        return false;
    }

    
    static bool ExecCompTimeOnly(InstrCompCellVal instrComp, TimeOnly cellValue, TimeOnly valueComp)
    {
        if (instrComp.Operator == InstrCompCellValOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == InstrCompCellValOperator.NotEqual)
            return cellValue != valueComp;

        if (instrComp.Operator == InstrCompCellValOperator.GreaterThan)
            return cellValue > valueComp;
        if (instrComp.Operator == InstrCompCellValOperator.GreaterOrEqualThan)
            return cellValue >= valueComp;

        if (instrComp.Operator == InstrCompCellValOperator.LesserThan)
            return cellValue < valueComp;
        if (instrComp.Operator == InstrCompCellValOperator.LesserOrEqualThan)
            return cellValue <= valueComp;
        return false;
    }

    static bool ExecCompString(InstrCompCellVal instrComp, string cellValue, string valueComp)
    {
        if (instrComp.Operator == InstrCompCellValOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == InstrCompCellValOperator.NotEqual)
            return cellValue != valueComp;


        return false;
    }

}
