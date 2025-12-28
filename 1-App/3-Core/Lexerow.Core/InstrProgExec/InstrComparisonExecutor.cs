using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Vml;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.Object;
using Lexerow.Core.Utils;
using OpenExcelSdk;

namespace Lexerow.Core.InstrProgExec;

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
    public bool ExecInstrComparison(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrComparison instrComparison)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrComparisonExecutor.ExecInstrComparison", string.Empty);

        InstrBase instrOperandLeft = instrComparison.OperandLeft;
        InstrBase instrOperandRight = instrComparison.OperandRight;

        //--a previous instr exists
        if (ctx.PrevInstrExecuted!=null)
        {
            // left instr executed before?
            if(instrComparison.LastInstrExecuted==1)
            {
                instrOperandLeft= ctx.PrevInstrExecuted;
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

        //--A.Cell > xxx 
        if (!CompareColCellFuncWith(result, _excelProcessor, ctx, progExecVarMgr, instrComparison, instrOperandLeft, instrOperandRight, out bool isCase))
            return false;
        if (isCase) return true;

        //--xxx<A.Cell, reverse both operands
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
            ProgExecVar progExecVar= progExecVarMgr.FindVarByName(instrNameObject.Name);
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
            if (!CompareColCellBlank(result, _excelProcessor, ctx.ExcelSheet, ctx.RowNum, instrColCellFuncLeft, instrComparison.Operator, out bool resultComp))
                return false;
            instrComparison.Result = resultComp;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        //--A.Cell=null or A.Cell<>null
        InstrNull instrNullRight = instrOperandRight as InstrNull;
        if (instrBlankRight != null)
        {
            isCase = true;
            if (!CompareColCellNull(result, _excelProcessor, ctx.ExcelSheet, ctx.RowNum, instrColCellFuncLeft, instrComparison.Operator, out bool resultComp))
                return false;
            instrComparison.Result = resultComp;
            ctx.PrevInstrExecuted = instrComparison;
            ctx.StackInstr.Pop();
            return true;
        }

        //--special case: A.Cell>B.Cell
        InstrColCellFunc instrColCellFuncRight = instrOperandRight as InstrColCellFunc;
        if (instrColCellFuncRight != null)
        {
            isCase = true;
            if (!Compare(result, _excelProcessor, ctx.ExcelFileObject.Filename, ctx.ExcelSheet, ctx.RowNum, instrColCellFuncLeft, instrComparison.Operator, instrColCellFuncRight, out bool resultComp))
                return false;

            instrComparison.Result = resultComp;
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
            if (!Compare(result, _excelProcessor, ctx.ExcelFileObject.Filename, ctx.ExcelSheet, ctx.RowNum, instrColCellFuncLeft, instrComparison.Operator, instrValueRight.ValueBase, out bool resultComp))
                return false;

            instrComparison.Result = resultComp;
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
    private bool Compare(Result result, ExcelProcessor excelProcessor, string fileName, ExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFuncLeft, InstrSepComparison compOperator, ValueBase valueRight, out bool resultComp)
    {
        resultComp = false;

        excelProcessor.GetCellTypeAndValue(excelSheet, instrColCellFuncLeft.ColNum, rowNum, out ExcelCell cell, out ExcelCellValueMulti excelCellValueMulti, out ExcelError error);

        // does the cell type match the If-Comparison cell.Value type?
        if (!ExcelExtendedUtils.MatchCellTypeAndIfComparison(excelCellValueMulti.CellType, valueRight))
        {
            // is there an warning already existing?
            // int sheetIndex= excelSheet.GetIndex()
            result.AddWarning(ErrorCode.ExecIfCondTypeMismatch, fileName, 0, instrColCellFuncLeft.ColNum, excelCellValueMulti.CellType);

            // just a warning, stop here but return true
            return true;
        }

        // execute the If part: comparison condition
        return CompareValues(result, excelProcessor, excelCellValueMulti, compOperator, valueRight, out resultComp);
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
    private bool Compare(Result result, ExcelProcessor excelProcessor, string fileName, ExcelSheet excelSheet, int rowNum, InstrColCellFunc instrColCellFuncLeft, InstrSepComparison compOperator, InstrColCellFunc instrColCellFuncRight, out bool resultComp)
    {
        resultComp = false;

        //var cellLeft = excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFuncLeft.ColNum - 1);
        excelProcessor.GetCellTypeAndValue(excelSheet, instrColCellFuncLeft.ColNum, rowNum, out ExcelCell cellLeft, out ExcelCellValueMulti excelCellValueMultiLeft, out ExcelError error);

        // get the cell value type
        //CellRawValueType cellTypeLeft = excelProcessor.GetCellValueType(excelSheet, cellLeft);

        //var cellRight = excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFuncRight.ColNum - 1);
        excelProcessor.GetCellTypeAndValue(excelSheet, instrColCellFuncRight.ColNum, rowNum, out ExcelCell cellRight, out ExcelCellValueMulti excelCellValueMultiRight, out ExcelError errorRight);

        // get the cell value type
        //CellRawValueType cellTypeRight = excelProcessor.GetCellValueType(excelSheet, cellRight);

        // does the cell type match the If-Comparison cell.Value type?
        if (!ExcelExtendedUtils.MatchCellTypeAndIfComparison(excelCellValueMultiLeft.CellType, excelCellValueMultiRight.CellType))
        {
            // is there an warning already existing?
            // int excelSheet.Index
            result.AddWarning(ErrorCode.ExecIfCondTypeMismatch, fileName, 0, instrColCellFuncLeft.ColNum, excelCellValueMultiLeft.CellType);
            // just a warning stop here but return true
            return true;
        }

        // execute the If part: comparison condition
        return CompareValues(result, excelProcessor, excelCellValueMultiLeft, compOperator, excelCellValueMultiRight, out resultComp);
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

        //var cell = excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFunc.ColNum - 1);
        excelProcessor.GetCellAt(excelSheet, instrColCellFunc.ColNum, rowNum, out ExcelCell cell, out ExcelError excelError);
        if (cell == null)
        {
            resultComp = true;
            return true;
        }

        // the cell exists but has no value
        if(string.IsNullOrEmpty(cell.Cell.InnerText))
        //string val = cell.GetRawValueString();
        //if (val == string.Empty)
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

        //var cell = excelProcessor.GetCellAt(excelSheet, rowNum, instrColCellFunc.ColNum - 1);
        excelProcessor.GetCellAt(excelSheet, instrColCellFunc.ColNum, rowNum, out ExcelCell cell, out ExcelError error);
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
    public static bool CompareValues(Result result, ExcelProcessor excelProcessor, ExcelCellValueMulti excelCellValueMulti, InstrSepComparison compOperator, ValueBase valueBase, out bool compResult)
    {
        //--integer
        if (excelCellValueMulti.CellType == ExcelCellType.Integer)
        {
            if (valueBase.ValueType == System.ValueType.Int)
            {
                compResult = CompareInt(excelCellValueMulti.IntegerValue.Value, compOperator, ((ValueInt)valueBase).Val);
                return true;
            }
            compResult = CompareDouble(excelCellValueMulti.IntegerValue.Value, compOperator, ((ValueDouble)valueBase).Val);
            return true;
        }

        //--double
        if (excelCellValueMulti.CellType == ExcelCellType.Double)
        {
            if (valueBase.ValueType == System.ValueType.Int)
            {
                compResult = CompareDouble(excelCellValueMulti.DoubleValue.Value, compOperator, ((ValueInt)valueBase).Val);
                return true;
            }
            compResult = CompareDouble(excelCellValueMulti.DoubleValue.Value, compOperator, ((ValueDouble)valueBase).Val);
            return true;
        }

        //--string
        if (excelCellValueMulti.CellType == ExcelCellType.String)
        {
            // only Equal or NotEqual is possible
            compResult = CompareString(excelCellValueMulti.StringValue, compOperator, ((ValueString)valueBase).Val);
            return true;
        }

        //--dateOnly
        if (excelCellValueMulti.CellType == ExcelCellType.DateOnly)
        {
            // dateOnly - dateOnly 
            if (valueBase.ValueType == System.ValueType.DateOnly)
            {
                compResult = CompareDateTime(excelCellValueMulti.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue), compOperator, ((ValueDateOnly)valueBase).ToDateTime());
                return true;
            }

            // dateOnly - dateTime
            compResult = CompareDateTime(excelCellValueMulti.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue), compOperator, ((ValueDateTime)valueBase).Val);
            return true;
        }

        //--dateTime
        if (excelCellValueMulti.CellType == ExcelCellType.DateTime)
        {
            // dateTime - dateOnly 
            if (valueBase.ValueType == System.ValueType.DateOnly)
            {
                compResult = CompareDateTime(excelCellValueMulti.DateTimeValue.Value, compOperator, ((ValueDateOnly)valueBase).ToDateTime());
                return true;
            }

            // dateTime - dateTime
            compResult = CompareDateTime(excelCellValueMulti.DateTimeValue.Value, compOperator, ((ValueDateTime)valueBase).Val);
            return true;
        }

        //--TimeOnly
        if (excelCellValueMulti.CellType == ExcelCellType.TimeOnly)
        {
            // should be a value less than 0, exp: 0,354746
            compResult = CompareTimeOnly(excelCellValueMulti.TimeOnlyValue.Value, compOperator, ((ValueTimeOnly)valueBase).Val);
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
    public static bool CompareValues(Result result, ExcelProcessor excelProcessor, ExcelCellValueMulti excelCellValueMultiLeft, InstrSepComparison compOperator, ExcelCellValueMulti excelCellValueMultiRight, out bool compResult)
    {
        //--left integer
        if (excelCellValueMultiLeft.CellType == ExcelCellType.Integer)
        {
            if (excelCellValueMultiRight.CellType == ExcelCellType.Integer)
            {
                compResult = CompareInt(excelCellValueMultiLeft.IntegerValue.Value, compOperator, excelCellValueMultiRight.IntegerValue.Value);
                return true;
            }

            if (excelCellValueMultiRight.CellType == ExcelCellType.Double)
            {
                compResult = CompareDouble(excelCellValueMultiLeft.IntegerValue.Value, compOperator, excelCellValueMultiRight.DoubleValue.Value);
                return true;
            }
        }

        //--left double
        if (excelCellValueMultiLeft.CellType == ExcelCellType.Double)
        {
            if (excelCellValueMultiRight.CellType == ExcelCellType.Double)
            {
                compResult = CompareDouble(excelCellValueMultiLeft.DoubleValue.Value, compOperator, excelCellValueMultiRight.DoubleValue.Value);
                return true;
            }

            if (excelCellValueMultiRight.CellType == ExcelCellType.Integer)
            {
                compResult = CompareDouble(excelCellValueMultiLeft.DoubleValue.Value, compOperator, excelCellValueMultiRight.IntegerValue.Value);
                return true;
            }
        }

        //--string
        if (excelCellValueMultiLeft.CellType == ExcelCellType.String)
        {
            compResult = CompareString(excelCellValueMultiLeft.StringValue, compOperator, excelCellValueMultiRight.StringValue);
            return true;
        }


        //--Left DateOnly
        if (excelCellValueMultiLeft.CellType == ExcelCellType.DateOnly)
        {
            if (excelCellValueMultiRight.CellType == ExcelCellType.DateOnly)
            {
                compResult = CompareDateTime(excelCellValueMultiLeft.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue), compOperator, excelCellValueMultiRight.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue));
                return true;
            }

            if (excelCellValueMultiRight.CellType == ExcelCellType.DateTime)
            {
                compResult = CompareDateTime(excelCellValueMultiLeft.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue), compOperator, excelCellValueMultiRight.DateTimeValue.Value);
                return true;
            }

        }

        //--Left DateTime
        if (excelCellValueMultiLeft.CellType == ExcelCellType.DateTime)
        {
            if (excelCellValueMultiRight.CellType == ExcelCellType.DateTime)
            {
                compResult = CompareDateTime(excelCellValueMultiLeft.DateTimeValue.Value, compOperator, excelCellValueMultiRight.DateTimeValue.Value);
                return true;
            }
            if (excelCellValueMultiRight.CellType == ExcelCellType.DateOnly)
            {
                compResult = CompareDateTime(excelCellValueMultiLeft.DateTimeValue.Value, compOperator, excelCellValueMultiRight.DateOnlyValue.Value.ToDateTime(TimeOnly.MinValue));
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

    private static bool CompareString(string cellValue, InstrSepComparison instrComp, string valueComp)
    {
        if (instrComp.Operator == SepComparisonOperator.Equal)
            return cellValue == valueComp;

        if (instrComp.Operator == SepComparisonOperator.Different)
            return cellValue != valueComp;

        return false;
    }
}