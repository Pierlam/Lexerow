using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.Utils;
using OpenExcelSdk;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Execute comparison instructions.
/// exp: A.Cell>10
/// </summary>
public class InstrComparisonExecutor
{
    private IActivityLogger _logger;

    private ExcelProcessor _excelProcessor;

    public InstrComparisonExecutor(IActivityLogger activityLogger, ExcelProcessor excelProcessor)
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
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrComparison"></param>
    /// <returns></returns>
    public bool ExecInstrCompExpr(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrComparison instrComparison)
    {
        _logger.LogExecStart(ActivityLogLevel.Debug, "InstrComparisonExecutor.ExecInstrComparison", string.Empty);

        InstrBase instrOperandLeft = instrComparison.OperandLeft;
        InstrBase instrOperandRight = instrComparison.OperandRight;

        //--a previous instr exists
        if (ctx.PrevInstrExecuted != null)
        {
            // left instr executed before?
            if (instrComparison.LastInstrExecuted == 1)
            {
                instrOperandLeft = ctx.PrevInstrExecuted;
                instrComparison.LastInstrExecuted = 0;
                ctx.PrevInstrExecuted = null;
            }
            // right instr executed before?
            if (instrComparison.LastInstrExecuted == 2)
            {
                instrOperandRight = ctx.PrevInstrExecuted;
                instrComparison.LastInstrExecuted = 0;
                ctx.PrevInstrExecuted = null;
            }
        }

        // the left operand is a function call, so execute it and come back here
        if (InstrUtils.NeedToBeExecuted(instrOperandLeft))
        {
            ctx.StackInstr.Push(instrOperandLeft);
            instrComparison.LastInstrExecuted = 1;
            return true;
        }

        // the right operand is a function call, so execute it and come back here
        if (InstrUtils.NeedToBeExecuted(instrOperandRight))
        {
            ctx.StackInstr.Push(instrOperandRight);
            instrComparison.LastInstrExecuted = 2;
            return true;
        }

        //--A.Cell > value
        if (!CompareColCellFuncWith(result, _excelProcessor, ctx, progExecVarMgr, instrComparison, instrOperandLeft, instrOperandRight, out bool isCase))
            return false;
        if (isCase) return true;

        //--value < A.Cell, reverse both operands
        InstrColCellFunc instrColCellFuncRight = instrOperandRight as InstrColCellFunc;
        if (instrColCellFuncRight != null)
        {
            InstrSepComparison sepCompRevert = instrComparison.Operator.Revert();
            if (!CompareColCellFuncWith(result, _excelProcessor, ctx, progExecVarMgr, instrComparison, instrOperandRight, instrOperandLeft, out isCase))
                return false;
            if (isCase) return true;
        }

        //--a=b
        // TODO:

        // case not managed
        result.AddError(ErrorCode.ExecInstrNotManaged, instrComparison.OperandLeft.FirstScriptToken());
        return false;
    }

    /// <summary>
    /// A.Cell>xxx
    /// Compare InstrColCellFucn to something: value, var, blank, null, Date.
    /// manage also the special case: A.Cell>B.Cell
    /// </summary>
    /// <param name="result"></param>
    /// <param name="excelProcessor"></param>
    /// <param name="ctx"></param>
    /// <param name="instrComparison"></param>
    /// <param name="instrOperandLeft"></param>
    /// <param name="instrOperandRight"></param>
    /// <param name="isCase"></param>
    /// <returns></returns>
    private bool CompareColCellFuncWith(Result result, ExcelProcessor excelProcessor, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrComparison instrComparison, InstrBase instrOperandLeft, InstrBase instrOperandRight, out bool isCase)
    {
        isCase = false;

        InstrColCellFunc instrColCellFuncLeft = instrOperandLeft as InstrColCellFunc;
        if (instrColCellFuncLeft == null) return true;

        //--right operand is a var?
        InstrNameObject instrNameObject = instrOperandRight as InstrNameObject;
        if (instrNameObject != null)
        {
            // get the value of the var, the inner one if the value is a var
            ProgExecVar progExecVar = progExecVarMgr.FindVarByName(instrNameObject.Name);
            if (progExecVar == null)
            {
                result.AddError(ErrorCode.ExecInstrVarNotFound, instrOperandRight.FirstScriptToken());
                return false;
            }
            // update the right operand with the (last) value of the var
            instrOperandRight = progExecVar.Value;
        }

        //--A.Cell=blank or A.Cell<>blank
        InstrBlank instrBlankRight = instrOperandRight as InstrBlank;
        if (instrBlankRight != null)
        {
            isCase = true;
            if (!CompareColCellBlank(result, _excelProcessor, ctx.ExcelSheet, ctx.RowAddr, instrColCellFuncLeft, instrComparison.Operator, out bool resultComp))
                return false;
            instrComparison.Result = resultComp;
            instrComparison.IsExecuted = true;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        //--A.Cell=null or A.Cell<>null
        InstrNull instrNullRight = instrOperandRight as InstrNull;
        if (instrBlankRight != null)
        {
            isCase = true;
            if (!CompareColCellNull(result, _excelProcessor, ctx.ExcelSheet, ctx.RowAddr, instrColCellFuncLeft, instrComparison.Operator, out bool resultComp))
                return false;
            instrComparison.Result = resultComp;
            instrComparison.IsExecuted = true;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        //--special case: A.Cell>B.Cell
        InstrColCellFunc instrColCellFuncRight = instrOperandRight as InstrColCellFunc;
        if (instrColCellFuncRight != null)
        {
            isCase = true;
            if (!Compare(result, _excelProcessor, progExecVarMgr, ctx.ExcelFileObject.Filename, ctx.ExcelSheet, ctx.RowAddr, instrColCellFuncLeft, instrComparison.Operator, instrColCellFuncRight, out bool resultComp))
                return false;

            instrComparison.Result = resultComp;
            instrComparison.IsExecuted = true;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        // right operand is basic value?
        InstrValue instrValueRight = instrOperandRight as InstrValue;

        //--A.Cell> value
        if (instrValueRight != null)
        {
            isCase = true;
            if (!Compare(result, _excelProcessor, progExecVarMgr, ctx.ExcelFileObject.Filename, ctx.ExcelSheet, ctx.RowAddr, instrColCellFuncLeft, instrComparison.Operator, instrValueRight.ValueBase, out bool resultComp))
                return false;

            instrComparison.Result = resultComp;
            instrComparison.IsExecuted = true;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        // not the case
        isCase = false;
        return true;
    }

    /// <summary>
    /// Compare InstrColCellFunc operator constValue.
    /// exp: A.Cell>10
    /// </summary>
    /// <param name="result"></param>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="instrColCellFuncLeft"></param>
    /// <param name="compOperator"></param>
    /// <param name="instrValueRight"></param>
    /// <param name="resultComp"></param>
    /// <returns></returns>
    /// <exception cref="NotImplementedException"></exception>
    private bool Compare(Result result, ExcelProcessor excelProcessor, ProgExecVarMgr progExecVarMgr, string fileName, ExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFuncLeft, InstrSepComparison compOperator, ValueBase valueRight, out bool resultComp)
    {
        resultComp = false;

        ExcelCellValue excelCellValue = excelProcessor.GetCellValue(excelSheet, instrColCellFuncLeft.ColNum, rowNum);

        // does the cell type match the If-Comparison cell.Value type?
        if (!ExcelExtendedUtils.MatchCellTypeAndIfComparison(excelCellValue.CellType, valueRight))
        {
            // is there an warning already existing?
            result.AddWarning(ErrorCode.ExecIfCondTypeMismatch, fileName, excelSheet.Index, instrColCellFuncLeft.ColNum, excelCellValue.CellType);

            // just a warning, stop here but return true
            return true;
        }

        // execute the If part: comparison condition
        return CompareValues(result, excelProcessor, progExecVarMgr, excelCellValue, compOperator, valueRight, out resultComp);
    }

    /// <summary>
    /// Compare InstrColCellFunc operator InstrColCellFunc
    /// exp: A.Cell>B.Cell
    /// </summary>
    /// <param name="result"></param>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="instrColCellFuncLeft"></param>
    /// <param name="compOperator"></param>
    /// <param name="instrValueRight"></param>
    /// <param name="resultComp"></param>
    /// <returns></returns>
    /// <exception cref="NotImplementedException"></exception>
    private bool Compare(Result result, ExcelProcessor excelProcessor, ProgExecVarMgr progExecVarMgr, string fileName, ExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFuncLeft, InstrSepComparison compOperator, InstrColCellFunc instrColCellFuncRight, out bool resultComp)
    {
        resultComp = false;

        ExcelCellValue excelCellValueLeft = excelProcessor.GetCellValue(excelSheet, instrColCellFuncLeft.ColNum, rowNum);
        ExcelCellValue excelCellValueRight = excelProcessor.GetCellValue(excelSheet, instrColCellFuncRight.ColNum, rowNum);


        // does the cell type match the If-Comparison cell.Value type?
        if (!ExcelExtendedUtils.MatchCellTypeAndIfComparison(excelCellValueLeft.CellType, excelCellValueRight.CellType))
        {
            // is there an warning already existing?
            result.AddWarning(ErrorCode.ExecIfCondTypeMismatch, fileName, excelSheet.Index, instrColCellFuncLeft.ColNum, excelCellValueLeft.CellType);
            // just a warning stop here but return true
            return true;
        }

        // execute the If part: comparison condition
        return CompareValues(result, excelProcessor, progExecVarMgr, excelCellValueLeft, compOperator, excelCellValueRight, out resultComp);
    }

    /// <summary>
    /// resultComp is true if the cell is blank (no value but can have a formating)
    /// or is the cell is null.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="excelProcessor"></param>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="instrColCellFunc"></param>
    /// <param name="compOperator"></param>
    /// <param name="resultComp"></param>
    /// <returns></returns>
    private static bool CompareColCellBlank(Result result, ExcelProcessor excelProcessor, ExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFunc, InstrSepComparison compOperator, out bool resultComp)
    {
        resultComp = false;

        ExcelCell cell = excelProcessor.GetCellAt(excelSheet, instrColCellFunc.ColNum, rowNum);
        if (cell == null)
        {
            resultComp = true;
            return true;
        }

        // the cell exists but has no value
        if (string.IsNullOrEmpty(cell.Cell.InnerText))
            resultComp = true;

        return true;
    }

    /// <summary>
    /// resultComp is true if the cell is null.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="excelProcessor"></param>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="instrColCellFunc"></param>
    /// <param name="compOperator"></param>
    /// <param name="resultComp"></param>
    /// <returns></returns>
    private static bool CompareColCellNull(Result result, ExcelProcessor excelProcessor, ExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFunc, InstrSepComparison compOperator, out bool resultComp)
    {
        resultComp = false;

        ExcelCell cell = excelProcessor.GetCellAt(excelSheet, instrColCellFunc.ColNum, rowNum);
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
    public static bool CompareValues(Result result, ExcelProcessor excelProcessor, ProgExecVarMgr progExecVarMgr, ExcelCellValue excelCellValue, InstrSepComparison compOperator, ValueBase valueBase, out bool compResult)
    {
        //--integer
        if (excelCellValue.CellType == ExcelCellType.Integer)
        {
            if (valueBase.ValueType == System.ValueType.Int)
            {
                compResult = CompareInt(excelCellValue.IntegerValue.Value, compOperator, ((ValueInt)valueBase).Val);
                return true;
            }
            compResult = CompareDouble(excelCellValue.IntegerValue.Value, compOperator, ((ValueDouble)valueBase).Val);
            return true;
        }

        //--double
        if (excelCellValue.CellType == ExcelCellType.Double)
        {
            if (valueBase.ValueType == System.ValueType.Int)
            {
                compResult = CompareDouble(excelCellValue.DoubleValue.Value, compOperator, ((ValueInt)valueBase).Val);
                return true;
            }
            compResult = CompareDouble(excelCellValue.DoubleValue.Value, compOperator, ((ValueDouble)valueBase).Val);
            return true;
        }

        //--string
        if (excelCellValue.CellType == ExcelCellType.String)
        {
            bool caseSensitive= progExecVarMgr.GetProgExecSysVarAsBool(CoreDefinitions.SysVarStrCompareCaseSensitive);

            // only Equal or NotEqual is possible
            compResult = CompareString(caseSensitive, excelCellValue.StringValue, compOperator, ((ValueString)valueBase).Val);
            return true;
        }

        //--dateOnly
        if (excelCellValue.CellType == ExcelCellType.DateOnly)
        {
            // dateOnly - dateOnly
            if (valueBase.ValueType == System.ValueType.DateOnly)
            {
                compResult = CompareDateTime(excelCellValue.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue), compOperator, ((ValueDateOnly)valueBase).ToDateTime());
                return true;
            }

            // dateOnly - dateTime
            compResult = CompareDateTime(excelCellValue.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue), compOperator, ((ValueDateTime)valueBase).Val);
            return true;
        }

        //--dateTime
        if (excelCellValue.CellType == ExcelCellType.DateTime)
        {
            // dateTime - dateOnly
            if (valueBase.ValueType == System.ValueType.DateOnly)
            {
                compResult = CompareDateTime(excelCellValue.DateTimeValue.Value, compOperator, ((ValueDateOnly)valueBase).ToDateTime());
                return true;
            }

            // dateTime - dateTime
            compResult = CompareDateTime(excelCellValue.DateTimeValue.Value, compOperator, ((ValueDateTime)valueBase).Val);
            return true;
        }

        //--TimeOnly
        if (excelCellValue.CellType == ExcelCellType.TimeOnly)
        {
            // should be a value less than 0, exp: 0,354746
            compResult = CompareTimeOnly(excelCellValue.TimeOnlyValue.Value, compOperator, ((ValueTimeOnly)valueBase).Val);
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
    public static bool CompareValues(Result result, ExcelProcessor excelProcessor, ProgExecVarMgr progExecVarMgr, ExcelCellValue excelCellValueLeft, InstrSepComparison compOperator, ExcelCellValue excelCellValueRight, out bool compResult)
    {
        //--left integer
        if (excelCellValueLeft.CellType == ExcelCellType.Integer)
        {
            if (excelCellValueRight.CellType == ExcelCellType.Integer)
            {
                compResult = CompareInt(excelCellValueLeft.IntegerValue.Value, compOperator, excelCellValueRight.IntegerValue.Value);
                return true;
            }

            if (excelCellValueRight.CellType == ExcelCellType.Double)
            {
                compResult = CompareDouble(excelCellValueLeft.IntegerValue.Value, compOperator, excelCellValueRight.DoubleValue.Value);
                return true;
            }
        }

        //--left double
        if (excelCellValueLeft.CellType == ExcelCellType.Double)
        {
            if (excelCellValueRight.CellType == ExcelCellType.Double)
            {
                compResult = CompareDouble(excelCellValueLeft.DoubleValue.Value, compOperator, excelCellValueRight.DoubleValue.Value);
                return true;
            }

            if (excelCellValueRight.CellType == ExcelCellType.Integer)
            {
                compResult = CompareDouble(excelCellValueLeft.DoubleValue.Value, compOperator, excelCellValueRight.IntegerValue.Value);
                return true;
            }
        }

        //--string
        if (excelCellValueLeft.CellType == ExcelCellType.String)
        {
            bool caseSensitive = progExecVarMgr.GetProgExecSysVarAsBool(CoreDefinitions.SysVarStrCompareCaseSensitive);
            compResult = CompareString(caseSensitive, excelCellValueLeft.StringValue, compOperator, excelCellValueRight.StringValue);
            return true;
        }

        //--Left DateOnly
        if (excelCellValueLeft.CellType == ExcelCellType.DateOnly)
        {
            if (excelCellValueRight.CellType == ExcelCellType.DateOnly)
            {
                compResult = CompareDateTime(excelCellValueLeft.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue), compOperator, excelCellValueRight.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue));
                return true;
            }

            if (excelCellValueRight.CellType == ExcelCellType.DateTime)
            {
                compResult = CompareDateTime(excelCellValueLeft.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue), compOperator, excelCellValueRight.DateTimeValue.Value);
                return true;
            }
        }

        //--Left DateTime
        if (excelCellValueLeft.CellType == ExcelCellType.DateTime)
        {
            if (excelCellValueRight.CellType == ExcelCellType.DateTime)
            {
                compResult = CompareDateTime(excelCellValueLeft.DateTimeValue.Value, compOperator, excelCellValueRight.DateTimeValue.Value);
                return true;
            }
            if (excelCellValueRight.CellType == ExcelCellType.DateOnly)
            {
                compResult = CompareDateTime(excelCellValueLeft.DateTimeValue.Value, compOperator, excelCellValueRight.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue));
                return true;
            }
        }

        compResult = false;
        return true;
    }

    /// <summary>
    /// CellValue operator value
    /// exp 10 >12
    /// </summary>
    /// <param name="instrComp"></param>
    /// <param name="cellValue"></param>
    /// <param name="valueComp"></param>
    /// <returns></returns>
    private static bool CompareInt(int cellValue, InstrSepComparison instrComp, int valueComp)
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

    /// <summary>
    /// CellValue operator value
    /// exp 23.5 > 12.89
    /// </summary>
    /// <param name="instrComp"></param>
    /// <param name="cellValue"></param>
    /// <param name="valueComp"></param>
    /// <returns></returns>
    private static bool CompareDouble(double cellValue, InstrSepComparison instrComp, double valueComp)
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

    private static bool CompareString(bool caseSensitive, string cellValue, InstrSepComparison instrComp, string valueComp)
    {
        if (instrComp.Operator == SepComparisonOperator.Equal)
        {
            // case sensitive
            if(caseSensitive) return cellValue.Equals(valueComp, StringComparison.InvariantCulture); 
            // case insensitive
            return cellValue.Equals(valueComp,StringComparison.InvariantCultureIgnoreCase);
        }
            

        if (instrComp.Operator == SepComparisonOperator.Different)
        {
            // case sensitive
            if (caseSensitive) return !cellValue.Equals(valueComp, StringComparison.InvariantCulture);
            // case insensitive
            return !cellValue.Equals(valueComp, StringComparison.InvariantCultureIgnoreCase);

        }
        return false;
    }
}