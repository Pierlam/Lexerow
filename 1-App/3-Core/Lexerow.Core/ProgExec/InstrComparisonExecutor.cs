using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.Utils;

namespace Lexerow.Core.ProgExec;

public class InstrComparisonExecutor
{
    private IActivityLogger _logger;

    private IExcelProcessor _excelProcessor;

    public InstrComparisonExecutor(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
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
    public bool ExecInstrComparison(ExecResult execResult, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrComparison instrComparison)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrComparisonRunner.RunInstrComparison", string.Empty);
        bool res;

        InstrColCellFunc instrColCellFuncLeft = instrComparison.OperandLeft as InstrColCellFunc;
        InstrValue instrConstValueRight = instrComparison.OperandRight as InstrValue;

        //--A.Cell>10
        if (instrColCellFuncLeft != null && instrConstValueRight != null)
        {
            if (!Compare(execResult, _excelProcessor, ctx.ExcelFileObject.Filename, ctx.ExcelSheet, ctx.RowNum, instrColCellFuncLeft, instrComparison.Operator, instrConstValueRight, out bool resultComp))
                return false;

            instrComparison.Result = resultComp;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        InstrValue instrConstValueLeft = instrComparison.OperandLeft as InstrValue;
        InstrColCellFunc instrColCellFuncRight = instrComparison.OperandRight as InstrColCellFunc;

        //--10<A.Cell
        if (instrConstValueLeft != null && instrColCellFuncRight != null)
        {
            // revert the operator,exp: < becomes >
            InstrSepComparison sepCompRevert = instrComparison.Operator.Revert();
            if (!Compare(execResult, _excelProcessor, ctx.ExcelFileObject.Filename, ctx.ExcelSheet, ctx.RowNum, instrColCellFuncLeft, sepCompRevert, instrConstValueRight, out bool resultComp))
                return false;

            instrComparison.Result = resultComp;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        //--A.Cell>B.Cell
        if (instrColCellFuncLeft != null && instrColCellFuncRight != null)
        {
            if (!Compare(execResult, _excelProcessor, ctx.ExcelFileObject.Filename, ctx.ExcelSheet, ctx.RowNum, instrColCellFuncLeft, instrComparison.Operator, instrColCellFuncRight, out bool resultComp))
                return false;

            instrComparison.Result = resultComp;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        //--A.Cell=blank  or A.Cell<>blank
        InstrBlank instrBlankRight = instrComparison.OperandRight as InstrBlank;
        if (instrColCellFuncLeft != null && instrBlankRight != null)
        {
            if (!CompareColCellBlank(execResult, _excelProcessor, ctx.ExcelSheet, ctx.RowNum, instrColCellFuncLeft, instrComparison.Operator, out bool resultComp))
                return false;
            instrComparison.Result = resultComp;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        //--A.Cell=null  or A.Cell<>null
        InstrNull instrNullRight = instrComparison.OperandRight as InstrNull;
        if (instrColCellFuncLeft != null && instrBlankRight != null)
        {
            if (!CompareColCellNull(execResult, _excelProcessor, ctx.ExcelSheet, ctx.RowNum, instrColCellFuncLeft, instrComparison.Operator, out bool resultComp))
                return false;
            instrComparison.Result = resultComp;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        //--A.Cell>val
        // TODO:

        //--a=b
        // TODO:

        execResult.AddError(ErrorCode.ExecInstrNotManaged, instrComparison.OperandLeft.FirstScriptToken());
        return false;
    }

    /// <summary>
    /// Compare InstrColCellFunc operator constValue
    /// exp: A.Cell>10
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="instrColCellFuncLeft"></param>
    /// <param name="compOperator"></param>
    /// <param name="instrConstValueRight"></param>
    /// <param name="resultComp"></param>
    /// <returns></returns>
    /// <exception cref="NotImplementedException"></exception>
    private bool Compare(ExecResult execResult, IExcelProcessor excelProcessor, string fileName, IExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFuncLeft, InstrSepComparison compOperator, InstrValue instrConstValueRight, out bool resultComp)
    {
        resultComp = false;

        var cell = excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFuncLeft.ColNum - 1);

        // get the cell value type
        CellRawValueType cellType = excelProcessor.GetCellValueType(excelSheet, cell);

        // does the cell type match the If-Comparison cell.Value type?
        if (!ExcelExtendedUtils.MatchCellTypeAndIfComparison(cellType, instrConstValueRight.ValueBase))
        {
            // is there an warning already existing?
            execResult.AddWarning(ErrorCode.IfCondTypeMismatch, fileName, excelSheet.Index, instrColCellFuncLeft.ColNum, cellType);
            // just a warning stop here but return true
            return true;
        }

        // execute the If part: comparison condition
        return CompareValues(execResult, excelProcessor, cell, compOperator, instrConstValueRight.ValueBase, out resultComp);
    }

    /// <summary>
    /// Compare InstrColCellFunc operator InstrColCellFunc
    /// exp: A.Cell>B.Cell
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="instrColCellFuncLeft"></param>
    /// <param name="compOperator"></param>
    /// <param name="instrConstValueRight"></param>
    /// <param name="resultComp"></param>
    /// <returns></returns>
    /// <exception cref="NotImplementedException"></exception>
    private bool Compare(ExecResult execResult, IExcelProcessor excelProcessor, string fileName, IExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFuncLeft, InstrSepComparison compOperator, InstrColCellFunc instrColCellFuncRight, out bool resultComp)
    {
        resultComp = false;

        var cellLeft = excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFuncLeft.ColNum - 1);

        // get the cell value type
        CellRawValueType cellTypeLeft = excelProcessor.GetCellValueType(excelSheet, cellLeft);

        var cellRight = excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFuncRight.ColNum - 1);

        // get the cell value type
        CellRawValueType cellTypeRight = excelProcessor.GetCellValueType(excelSheet, cellRight);

        // does the cell type match the If-Comparison cell.Value type?
        if (!ExcelExtendedUtils.MatchCellTypeAndIfComparison(cellTypeLeft, cellTypeRight))
        {
            // is there an warning already existing?
            execResult.AddWarning(ErrorCode.IfCondTypeMismatch, fileName, excelSheet.Index, instrColCellFuncLeft.ColNum, cellTypeLeft);
            // just a warning stop here but return true
            return true;
        }

        // execute the If part: comparison condition
        return CompareValues(execResult, excelProcessor, cellLeft, cellTypeLeft, compOperator, cellRight, cellTypeRight, out resultComp);
    }

    /// <summary>
    /// resultComp is true if the cell is blank (no value but can have a formating)
    /// or is the cell is null.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="excelProcessor"></param>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="instrColCellFunc"></param>
    /// <param name="compOperator"></param>
    /// <param name="resultComp"></param>
    /// <returns></returns>
    private static bool CompareColCellBlank(ExecResult execResult, IExcelProcessor excelProcessor, IExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFunc, InstrSepComparison compOperator, out bool resultComp)
    {
        resultComp = false;

        var cell = excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFunc.ColNum - 1);
        if (cell == null)
        {
            resultComp = true;
            return true;
        }

        // the cell exists but has no value
        string val = cell.GetRawValueString();
        if (val == string.Empty)
            resultComp = true;

        return true;
    }

    /// <summary>
    /// resultComp is true if the cell is null.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="excelProcessor"></param>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="instrColCellFunc"></param>
    /// <param name="compOperator"></param>
    /// <param name="resultComp"></param>
    /// <returns></returns>
    private static bool CompareColCellNull(ExecResult execResult, IExcelProcessor excelProcessor, IExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFunc, InstrSepComparison compOperator, out bool resultComp)
    {
        resultComp = false;

        var cell = excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFunc.ColNum - 1);
        if (cell == null)
        {
            resultComp = true;
            return true;
        }

        // the cell exists, so it's not null
        resultComp = false;
        return true;
    }

    /// <summary>
    /// Execute the comparison condition.
    /// Type of cell value and instr comparison should match.
    ///
    /// If -condition- Then
    /// A.Cell>10
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="SheetNum"></param>
    /// <param name="listCols"></param>
    /// <param name="instrComp"></param>
    /// <param name="compResult"></param>
    /// <returns></returns>
    public static bool CompareValues(ExecResult execResult, IExcelProcessor excelProcessor, IExcelCell excelCellLeft, InstrSepComparison compOperator, ValueBase valueBase, out bool compResult)
    {
        if (valueBase.ValueType == System.ValueType.Int)
        {
            double dblVal = excelCellLeft.GetRawValueNumeric();
            compResult = CompareNumeric(dblVal, compOperator, ((ValueInt)valueBase).Val);
            return true;
        }

        if (valueBase.ValueType == System.ValueType.Double)
        {
            double dblVal = excelCellLeft.GetRawValueNumeric();
            compResult = CompareNumeric(dblVal, compOperator, ((ValueDouble)valueBase).Val);
            return true;
        }

        if (valueBase.ValueType == System.ValueType.String)
        {
            string stringVal = excelCellLeft.GetRawValueString();
            // only Equal or NotEqual is possible
            compResult = CompareString(stringVal, compOperator, ((ValueString)valueBase).Val);
            return true;
        }

        if (valueBase.ValueType == System.ValueType.DateOnly)
        {
            double doubleVal = excelCellLeft.GetRawValueNumeric();
            DateTime dtVal = DateTime.FromOADate(doubleVal);
            compResult = CompareDateTime(dtVal, compOperator, ((ValueDateOnly)valueBase).ToDateTime());
            return true;
        }

        if (valueBase.ValueType == System.ValueType.TimeOnly)
        {
            // should be a value less than 0, exp: 0,354746
            double doubleVal = excelCellLeft.GetRawValueNumeric();
            TimeOnly timeOnly = DateTimeUtils.ToTimeOnly(doubleVal);
            compResult = CompareTimeOnly(timeOnly, compOperator, ((ValueTimeOnly)valueBase).Val);
            return true;
        }

        if (valueBase.ValueType == System.ValueType.DateTime)
        {
            double doubleVal = excelCellLeft.GetRawValueNumeric();
            DateTime dtVal = DateTimeUtils.ToDateTime(doubleVal);
            compResult = CompareDateTime(dtVal, compOperator, ((ValueDateTime)valueBase).Val);
            return true;
        }

        compResult = false;
        return true;
    }

    /// <summary>
    /// Execute the comparison condition.
    /// Type of cell value and instr comparison should match.
    ///
    /// If -condition- Then
    /// A.Cell>B.Cell
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="SheetNum"></param>
    /// <param name="listCols"></param>
    /// <param name="instrComp"></param>
    /// <param name="compResult"></param>
    /// <returns></returns>
    public static bool CompareValues(ExecResult execResult, IExcelProcessor excelProcessor, IExcelCell excelCellLeft, CellRawValueType cellTypeLeft, InstrSepComparison compOperator, IExcelCell excelCellRight, CellRawValueType cellTypeRight, out bool compResult)
    {
        if (cellTypeLeft == CellRawValueType.Numeric)
        {
            double dblValLeft = excelCellLeft.GetRawValueNumeric();
            double dblValRight = excelCellRight.GetRawValueNumeric();
            compResult = CompareNumeric(dblValLeft, compOperator, dblValRight);
            return true;
        }

        if (cellTypeLeft == CellRawValueType.String)
        {
            string stringValLeft = excelCellLeft.GetRawValueString();
            string stringValRight = excelCellLeft.GetRawValueString();
            // only Equal or NotEqual is possible
            compResult = CompareString(stringValLeft, compOperator, stringValRight);
            return true;
        }

        if (cellTypeLeft == CellRawValueType.DateOnly && cellTypeRight == CellRawValueType.DateOnly)
        {
            double doubleValLeft = excelCellLeft.GetRawValueNumeric();
            DateTime dtValLeft = DateTime.FromOADate(doubleValLeft);

            double doubleValRight = excelCellRight.GetRawValueNumeric();
            DateTime dtValRight = DateTime.FromOADate(doubleValRight);

            compResult = CompareDateTime(dtValLeft, compOperator, dtValRight);
            return true;
        }

        if (cellTypeLeft == CellRawValueType.TimeOnly && cellTypeRight == CellRawValueType.TimeOnly)
        {
            // should be a value less than 0, exp: 0,354746
            double doubleValLeft = excelCellLeft.GetRawValueNumeric();
            TimeOnly timeOnlyLeft = DateTimeUtils.ToTimeOnly(doubleValLeft);

            double doubleValRight = excelCellRight.GetRawValueNumeric();
            TimeOnly timeOnlyRight = DateTimeUtils.ToTimeOnly(doubleValRight);

            compResult = CompareTimeOnly(timeOnlyLeft, compOperator, timeOnlyRight);
            return true;
        }

        if (cellTypeLeft == CellRawValueType.DateTime && cellTypeRight == CellRawValueType.DateTime)
        {
            double doubleValLeft = excelCellLeft.GetRawValueNumeric();
            DateTime dtValLeft = DateTimeUtils.ToDateTime(doubleValLeft);

            double doubleValRight = excelCellRight.GetRawValueNumeric();
            DateTime dtValRight = DateTimeUtils.ToDateTime(doubleValRight);
            compResult = CompareDateTime(dtValLeft, compOperator, dtValRight);
            return true;
        }

        if (cellTypeLeft == CellRawValueType.DateTime && cellTypeRight == CellRawValueType.DateOnly)
        {
            double doubleValLeft = excelCellLeft.GetRawValueNumeric();
            DateTime dtValLeft = DateTimeUtils.ToDateTime(doubleValLeft);

            double doubleValRight = excelCellRight.GetRawValueNumeric();
            DateTime dtValRight = DateTime.FromOADate(doubleValRight);

            compResult = CompareDateTime(dtValLeft, compOperator, dtValRight);
            return true;
        }

        if (cellTypeLeft == CellRawValueType.DateOnly && cellTypeRight == CellRawValueType.DateTime)
        {
            double doubleValLeft = excelCellLeft.GetRawValueNumeric();
            DateTime dtValLeft = DateTime.FromOADate(doubleValLeft);

            double doubleValRight = excelCellRight.GetRawValueNumeric();
            DateTime dtValRight = DateTimeUtils.ToDateTime(doubleValRight);
            compResult = CompareDateTime(dtValLeft, compOperator, dtValRight);
            return true;
        }

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
    private static bool CompareNumeric(double cellValue, InstrSepComparison instrComp, double valueComp)
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

    private static bool CompareDateTime(DateTime cellValue, InstrSepComparison instrComp, DateTime valueComp)
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

    private static bool CompareTimeOnly(TimeOnly cellValue, InstrSepComparison instrComp, TimeOnly valueComp)
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

    private static bool CompareString(string cellValue, InstrSepComparison instrComp, string valueComp)
    {
        if (instrComp.Operator == SepComparisonOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == SepComparisonOperator.Different)
            return cellValue != valueComp;

        return false;
    }
}