using Lexerow.Core.Core.Exec;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.ProgRun;
using Lexerow.Core.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgRun;
public class InstrComparisonRunner
{
    IActivityLogger _logger;

    IExcelProcessor _excelProcessor;

    public InstrComparisonRunner(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
    }

    /// <summary>
    /// Run comparison.
    /// cases:
    ///   A.Cell=10
    ///   10=A.Cell
    ///   A.Cell=B.Cell
    ///   Fct()=12
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrComparison"></param>
    /// <returns></returns>
    public bool RunInstrComparison(ExecResult execResult, ProgramRunnerContext ctx, List<ExecVar> listVar, InstrComparison instrComparison)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "InstrComparisonRunner.RunInstrComparison", string.Empty);
        bool res;

        // is left instr A.Cell?
        InstrColCellFunc instrColCellFunc = instrComparison.OperandLeft as InstrColCellFunc;
        InstrConstValue instrConstValue= instrComparison.OperandRight as InstrConstValue;

        // A.Cell>10
        if (instrColCellFunc != null && instrConstValue != null)
        {
            if (!Compare(execResult, ctx.ExcelFileObject.Filename,  ctx.ExcelSheet, ctx.RowNum, instrColCellFunc, instrComparison.Operator, instrConstValue, out bool resultComp))
                return false;

            instrComparison.Result= resultComp;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        // 10<A.Cell

        // A.Cell>B.Cell


        return true;
    }

    /// <summary>
    /// Compare InstrColCellFunc operator constValue
    /// exp: A.Cell>10
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="instrColCellFunc"></param>
    /// <param name="compOperator"></param>
    /// <param name="instrConstValue"></param>
    /// <param name="resultComp"></param>
    /// <returns></returns>
    /// <exception cref="NotImplementedException"></exception>
    bool Compare(ExecResult execResult, string fileName,IExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFunc, InstrSepComparison compOperator, InstrConstValue instrConstValue, out bool resultComp)
    {
        resultComp = false;

        var cell = _excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFunc.ColNum-1);

        // get the cell value type            
        CellRawValueType cellType = _excelProcessor.GetCellValueType(excelSheet, cell);

        // does the cell type match the If-Comparison cell.Value type?
        if (!ExcelExtendedUtils.MatchCellTypeAndIfComparison(cellType, instrConstValue.ValueBase))
        {
            // is there an warning already existing? 
            execResult.AddWarning(ErrorCode.IfCondTypeMismatch, fileName, excelSheet.Index, instrColCellFunc.ColNum, cellType);
            // just a warning stop here but return true
            return true;
        }

        // execute the If part: comparison condition
        return CompareValues(execResult, _excelProcessor, cell, compOperator, instrConstValue.ValueBase, out resultComp);

    }

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
    public static bool CompareValues(ExecResult execResult, IExcelProcessor excelProcessor, IExcelCell excelCell, InstrSepComparison compOperator, ValueBase valueBase, out bool compResult)
    {
        if (valueBase.ValueType == System.ValueType.Int)
        {
            double dblVal = excelCell.GetRawValueNumeric();
            compResult = ExecCompNumeric(dblVal, compOperator, ((ValueInt)valueBase).Val);
            return true;
        }

        if (valueBase.ValueType == System.ValueType.Double)
        {
            double dblVal = excelCell.GetRawValueNumeric();
            compResult = ExecCompNumeric(dblVal, compOperator, ((ValueDouble)valueBase).Val);
            return true;
        }

        if (valueBase.ValueType == System.ValueType.String)
        {
            string stringVal = excelCell.GetRawValueString();
            // only Equal or NotEqual is possible
            compResult = ExecCompString(stringVal, compOperator, ((ValueString)valueBase).Val);
            return true;
        }

        if (valueBase.ValueType == System.ValueType.DateOnly)
        {
            double doubleVal = excelCell.GetRawValueNumeric();
            DateTime dtVal = DateTime.FromOADate(doubleVal);
            compResult = ExecCompDateTime(dtVal, compOperator, ((ValueDateOnly)valueBase).ToDateTime());
            return true;
        }

        if (valueBase.ValueType == System.ValueType.TimeOnly)
        {
            // should be a value less than 0, exp: 0,354746
            double doubleVal = excelCell.GetRawValueNumeric();
            TimeOnly timeOnly = DateTimeUtils.ToTimeOnly(doubleVal);
            compResult = ExecCompTimeOnly(timeOnly, compOperator, ((ValueTimeOnly)valueBase).Val);
            return true;
        }

        if (valueBase.ValueType == System.ValueType.DateTime)
        {
            double doubleVal = excelCell.GetRawValueNumeric();
            DateTime dtVal = DateTimeUtils.ToDateTime(doubleVal);
            compResult = ExecCompDateTime(dtVal, compOperator, ((ValueDateTime)valueBase).Val);
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
        if (excelCell == null || excelCell.GetRawValueString().Length == 0)
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
        if (stringVal.Length == 0) return false;

        StringComparison stringComparison = StringComparison.InvariantCulture;
        if (!instr.CaseSensitive)
            stringComparison = StringComparison.InvariantCultureIgnoreCase;

        foreach (string item in instr.ListItems)
        {
            if (stringVal.Equals(item, stringComparison)) return true;
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
    static bool ExecCompNumeric(double cellValue, InstrSepComparison instrComp, double valueComp)
    {
        if (instrComp.Operator == SepComparisonOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == SepComparisonOperator.Different)
            return cellValue != valueComp;

        if (instrComp.Operator == SepComparisonOperator.GreaterThan)
            return cellValue > valueComp;
        if (instrComp.Operator == SepComparisonOperator.GreaterEqualThan)
            return cellValue >= valueComp;

        if (instrComp.Operator == SepComparisonOperator.LessThan)
            return cellValue < valueComp;
        if (instrComp.Operator == SepComparisonOperator.LessEqualThan)
            return cellValue <= valueComp;
        return false;
    }

    static bool ExecCompDateTime(DateTime cellValue, InstrSepComparison instrComp, DateTime valueComp)
    {
        if (instrComp.Operator == SepComparisonOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == SepComparisonOperator.Different)
            return cellValue != valueComp;

        if (instrComp.Operator == SepComparisonOperator.GreaterThan)
            return cellValue > valueComp;
        if (instrComp.Operator == SepComparisonOperator.GreaterEqualThan)
            return cellValue >= valueComp;

        if (instrComp.Operator == SepComparisonOperator.LessThan)
            return cellValue < valueComp;
        if (instrComp.Operator == SepComparisonOperator.LessEqualThan)
            return cellValue <= valueComp;
        return false;
    }


    static bool ExecCompTimeOnly(TimeOnly cellValue, InstrSepComparison instrComp,  TimeOnly valueComp)
    {
        if (instrComp.Operator == SepComparisonOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == SepComparisonOperator.Different)
            return cellValue != valueComp;

        if (instrComp.Operator == SepComparisonOperator.GreaterThan)
            return cellValue > valueComp;
        if (instrComp.Operator == SepComparisonOperator.GreaterEqualThan)
            return cellValue >= valueComp;

        if (instrComp.Operator == SepComparisonOperator.LessThan)
            return cellValue < valueComp;
        if (instrComp.Operator == SepComparisonOperator.LessEqualThan)
            return cellValue <= valueComp;
        return false;
    }

    static bool ExecCompString(string cellValue, InstrSepComparison instrComp, string valueComp)
    {
        if (instrComp.Operator == SepComparisonOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == SepComparisonOperator.Different)
            return cellValue != valueComp;


        return false;
    }


}
