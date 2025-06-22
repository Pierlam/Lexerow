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
    public static bool ExecInstrCompCellVal(ExecResult execResult, IExcelProcessor excelProcessor, InstrCompColCellVal instrComp, IExcelCell excelCell, out bool compResult)
    {
        if (instrComp.Value.ValueType == System.ValueType.Int)
        {
            double dblVal = excelCell.GetRawValueNumeric();
            compResult= ExecCompNumeric(instrComp, dblVal, ((ValueInt)instrComp.Value).Val);
            return true;  
        }

        if (instrComp.Value.ValueType == System.ValueType.Double)
        {
            double dblVal = excelCell.GetRawValueNumeric();
            compResult = ExecCompNumeric(instrComp, dblVal, ((ValueDouble)instrComp.Value).Val);
            return true;
        }

        if (instrComp.Value.ValueType== System.ValueType.String)
        {
            string stringVal= excelCell.GetRawValueString();
            // only Equal or NotEqual is possible
            compResult = ExecCompString(instrComp, stringVal, ((ValueString)instrComp.Value).Val);
            return true;
        }

        if (instrComp.Value.ValueType == System.ValueType.DateOnly)
        {
            double doubleVal = excelCell.GetRawValueNumeric();
            DateTime dtVal = DateTime.FromOADate(doubleVal);
            compResult = ExecCompDateTime(instrComp, dtVal, ((ValueDateOnly)instrComp.Value).ToDateTime());
            return true;
        }

        if (instrComp.Value.ValueType == System.ValueType.TimeOnly)
        {
            // should be a value less than 0, exp: 0,354746
            double doubleVal = excelCell.GetRawValueNumeric();
            TimeOnly timeOnly= DateTimeUtils.ToTimeOnly(doubleVal);
            compResult = ExecCompTimeOnly(instrComp, timeOnly, ((ValueTimeOnly)instrComp.Value).Val);
            return true;
        }

        if (instrComp.Value.ValueType == System.ValueType.DateTime)
        {
            double doubleVal = excelCell.GetRawValueNumeric();
            DateTime dtVal = DateTimeUtils.ToDateTime(doubleVal);
            compResult = ExecCompDateTime(instrComp, dtVal, ((ValueDateTime)instrComp.Value).Val);
            return true;
        }

        compResult = false;
        return true;
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
    public static bool ExecInstrCompCellValIsNull(ExecResult execResult, IExcelProcessor excelProcessor, InstrCompColCellValIsNull instrComp, IExcelCell excelCell, out bool compResult)
    {
        if (excelCell == null || excelCell.GetRawValueString().Length==0)
        {
            // if A.Cell = null -> yes
            if (instrComp.Operator == ValCompOperator.Equal)
            {
                compResult = true;
                return true;
            }
            // if A.Cell = null -> no
            if (instrComp.Operator == ValCompOperator.NotEqual)
            {
                compResult = false;
                return true;
            }
            // XX ERROR
            compResult = false;
            return true;
        }

        // the cell hasa value
        // if A.Cell <> null -> yes
        if (instrComp.Operator == ValCompOperator.NotEqual)
        {
            compResult = true;
            return true;
        }

        // if A.Cell = null -> no
        if (instrComp.Operator == ValCompOperator.Equal)
        {
            compResult = false;
            return true;
        }

        // XX ERROR
        compResult = false;
            return true;

    }

    /// <summary>
    /// A.Cell In [ "y", "yes", "ok"]
    /// in string items.
    /// </summary>
    /// <param name="excelProcessor"></param>
    /// <param name="instr"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public static bool ExecInstrCompColCellInStringItems(IExcelProcessor excelProcessor, InstrCompColCellInStringItems instr, IExcelCell excelCell)
    {
        if (excelCell == null) return false;

        string stringVal = excelCell.GetRawValueString();
        if(stringVal.Length==0) return false;

        StringComparison stringComparison = StringComparison.InvariantCulture;
        if(!instr.CaseSensitive)
            stringComparison= StringComparison.InvariantCultureIgnoreCase;

        foreach (string item in instr.ListItems)
        {
            if(stringVal.Equals(item, stringComparison))return true;
        }

        return false;
    }

    /// <summary>
    /// CellValue operator value
    /// exp >12
    /// </summary>
    /// <param name="instrComp"></param>
    /// <param name="cellValue"></param>
    /// <param name="valueComp"></param>
    /// <returns></returns>
    static bool ExecCompNumeric(InstrCompColCellVal instrComp, double cellValue, double valueComp)
    {
        if(instrComp.Operator== ValCompOperator.Equal) 
            return cellValue == valueComp;

        if (instrComp.Operator == ValCompOperator.NotEqual)
            return cellValue != valueComp;

        if (instrComp.Operator == ValCompOperator.GreaterThan)
            return cellValue > valueComp;
        if (instrComp.Operator == ValCompOperator.GreaterOrEqualThan)
            return cellValue >= valueComp;

        if (instrComp.Operator == ValCompOperator.LessThan)
            return cellValue < valueComp;
        if (instrComp.Operator == ValCompOperator.LessOrEqualThan)
            return cellValue <= valueComp;
        return false;
    }

    static bool ExecCompDateTime(InstrCompColCellVal instrComp, DateTime cellValue, DateTime valueComp)
    {
        if (instrComp.Operator == ValCompOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == ValCompOperator.NotEqual)
            return cellValue != valueComp;

        if (instrComp.Operator == ValCompOperator.GreaterThan)
            return cellValue > valueComp;
        if (instrComp.Operator == ValCompOperator.GreaterOrEqualThan)
            return cellValue >= valueComp;

        if (instrComp.Operator == ValCompOperator.LessThan)
            return cellValue < valueComp;
        if (instrComp.Operator == ValCompOperator.LessOrEqualThan)
            return cellValue <= valueComp;
        return false;
    }

    
    static bool ExecCompTimeOnly(InstrCompColCellVal instrComp, TimeOnly cellValue, TimeOnly valueComp)
    {
        if (instrComp.Operator == ValCompOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == ValCompOperator.NotEqual)
            return cellValue != valueComp;

        if (instrComp.Operator == ValCompOperator.GreaterThan)
            return cellValue > valueComp;
        if (instrComp.Operator == ValCompOperator.GreaterOrEqualThan)
            return cellValue >= valueComp;

        if (instrComp.Operator == ValCompOperator.LessThan)
            return cellValue < valueComp;
        if (instrComp.Operator == ValCompOperator.LessOrEqualThan)
            return cellValue <= valueComp;
        return false;
    }

    static bool ExecCompString(InstrCompColCellVal instrComp, string cellValue, string valueComp)
    {
        if (instrComp.Operator == ValCompOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == ValCompOperator.NotEqual)
            return cellValue != valueComp;


        return false;
    }

}
