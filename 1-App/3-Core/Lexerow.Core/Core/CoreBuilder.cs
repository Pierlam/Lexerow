using Lexerow.Core.System;
using Lexerow.Core.Utils;
using Microsoft.Extensions.Logging;
using NPOI.HPSF;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core;
public class CoreBuilder
{
    ILoggerFactory _loggerFactory;

    CoreData _coreData;

    public CoreBuilder(ILoggerFactory loggerFactory, CoreData coreData)
    {
        _loggerFactory = loggerFactory;
        _coreData = coreData;
    }

    /// <summary>
    /// Create an instruction, Cell Value comparison.
    /// exp: CellVal > 12
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompCellVal CreateInstrCompCellVal(int colNum, InstrCompCellValOperator oper, int value)
    {
        // convert the value to an object Value
        ValueInt valueInt = new ValueInt(value);

        InstrCompCellVal exprComp = new InstrCompCellVal(colNum, oper, valueInt);

        return exprComp;
    }

    /// <summary>
    /// Create an instruction, Cell Value comparison.
    /// exp: CellVal > 12,3
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompCellVal CreateInstrCompCellVal(int colNum, InstrCompCellValOperator oper, double value)
    {
        // convert the value to an object Value
        ValueDouble valueDouble = new ValueDouble(value);

        InstrCompCellVal exprComp = new InstrCompCellVal(colNum, oper, valueDouble);

        return exprComp;
    }

    /// <summary>
    /// Create a Cell Value expression comparison.
    /// exp: CellVal = "georges" 
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompCellVal CreateInstrCompCellVal(int colNum, InstrCompCellValOperator oper, string value)
    {
        // convert the value to an object Value
        ValueString valueString = new ValueString(value);

        InstrCompCellVal exprComp = new InstrCompCellVal(colNum, oper, valueString);

        return exprComp;
    }

    /// <summary>
    /// Create an instruction, Cell Value comparison.
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompCellVal CreateInstrCompCellVal(int colNum, InstrCompCellValOperator oper, DateOnly value)
    {
        // convert the value to an object Value
        ValueDateOnly valueDateOnly = new ValueDateOnly(value);

        InstrCompCellVal exprComp = new InstrCompCellVal(colNum, oper, valueDateOnly);

        return exprComp;
    }

    /// <summary>
    /// Create an instruction, Cell Value comparison.
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompCellVal CreateInstrCompCellVal(int colNum, InstrCompCellValOperator oper, DateTime value)
    {
        // convert the value to an object Value
        ValueDateTime valueDateTime = new ValueDateTime(value);

        InstrCompCellVal exprComp = new InstrCompCellVal(colNum, oper, valueDateTime);

        return exprComp;
    }

    /// <summary>
    /// Create an instruction, Cell Value comparison.
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompCellVal CreateInstrCompCellVal(int colNum, InstrCompCellValOperator oper, TimeOnly value)
    {
        // convert the value to an object Value
        ValueTimeOnly valueTimeOnly = new ValueTimeOnly(value);

        InstrCompCellVal exprComp = new InstrCompCellVal(colNum, oper, valueTimeOnly);

        return exprComp;
    }

    /// <summary>
    /// Create Comp instr: B.Cell=Null.
    /// </summary>
    /// <param name="colNum"></param>
    /// <param name="oper"></param>
    /// <returns></returns>
    public InstrCompCellValIsNull CreateInstrCompCellValIsNull(int colNum, InstrCompCellValOperator oper)
    {
        return new InstrCompCellValIsNull(colNum, oper);
    }

    /// <summary>
    /// Create an instruction Set CellValue := val
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrSetCellVal CreateInstrSetCellVal(int colNum, double value)
    {
        ValueDouble valueDouble = new ValueDouble(value);

        return new InstrSetCellVal(colNum, valueDouble);
    }

    /// <summary>
    /// Create an instruction Set CellValue := val
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrSetCellVal CreateInstrSetCellVal(int colNum, string value)
    {
        ValueString valueString = new ValueString(value);

        return new InstrSetCellVal(colNum, valueString);
    }

    /// <summary>
    /// Create an instruction Set CellValue := val
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrSetCellVal CreateInstrSetCellVal(int colNum, DateOnly value)
    {
        ValueDateOnly valueDateOnly = new ValueDateOnly(value);

        return new InstrSetCellVal(colNum, valueDateOnly);
    }

    /// <summary>
    /// Create an instruction Set CellValue := val
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrSetCellVal CreateInstrSetCellVal(int colNum, DateTime value)
    {
        ValueDateTime valueDateTime = new ValueDateTime(value);

        return new InstrSetCellVal(colNum, valueDateTime);
    }

    /// <summary>
    /// Create an instruction Set CellValue := val
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrSetCellVal CreateInstrSetCellVal(int colNum, TimeOnly value)
    {
        ValueTimeOnly valueTimeOnly = new ValueTimeOnly(value);

        return new InstrSetCellVal(colNum, valueTimeOnly);
    }

    public InstrSetCellNull CreateInstrSetCellNull(int colNum)
    {
        return new InstrSetCellNull(colNum);
    }
    public InstrSetCellValueBlank CreateInstrSetCellValueBlank(int colNum)
    {
        return new InstrSetCellValueBlank(colNum);
    }

    /// <summary>
    /// Create an instruction OpenExcel.
    /// exp: file=OpenExcel(fileName)
    /// </summary>
    /// <param name="fileObjectRet"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public ExecResult CreateInstrOpenExcel(string excelFileObjectName, string fileName)
    {
        ExecResult execResult = new ExecResult();

        // possible to create the instr?
        if(_coreData.Stage != CoreStage.Build)
        {
            execResult.AddError(new CoreError(ErrorCode.UnableCreateInstrNotInStageBuild, null));
            return execResult;
        }

        if (string.IsNullOrWhiteSpace(excelFileObjectName))
        {
            execResult.AddError(new CoreError(ErrorCode.FileObjectNameNullOrEmpty, null));
            return execResult;
        }

        if (string.IsNullOrWhiteSpace(fileName))
        {
            execResult.AddError(new CoreError(ErrorCode.FileNameNullOrEmpty, null));
            return execResult;
        }

        // check the syntax of the excel file object name
        if (!SyntaxUtils.CheckIdSyntax(excelFileObjectName)) 
        {
            execResult.AddError(new CoreError(ErrorCode.ExcelFileObjectNameSyntaxWrong, null));
            return execResult;
        }

        // ok create the instruction
        InstrOpenExcel instrOpenExcel = new InstrOpenExcel(excelFileObjectName, fileName);

        // check the excel filename, should be defined previsouly
        if (FindExcelFileName(fileName))
        {
            // todo: fix code error
            execResult.AddError(new CoreError(ErrorCode.ExcelFileNameAlreadyOpen, null));
            return execResult;
        }

        // check the excelFileObject name, should be defined previsouly
        if (FindExcelFileObjectName(excelFileObjectName))
        {
            // todo: fix code error
            execResult.AddError(new CoreError(ErrorCode.ExcelFileObjectNameAlreadyOpen, null));
            return execResult;
        }

        _coreData.ListInstr.Add(instrOpenExcel);
        return execResult;
    }

    public ExecResult CreateInstrForEachRowIfThen(string excelFileObjectName, int sheetNum, int firstDataRowNumline, InstrBase instrIf, InstrBase instrThen)
    {
        List<InstrBase> listInstr= new List<InstrBase>
        {
            instrThen
        };
        return CreateInstrForEachRowIfThen(excelFileObjectName, sheetNum, firstDataRowNumline, instrIf, listInstr);
    }

    public ExecResult CreateInstrForEachRowIfThen(string excelFileObjectName, int sheetNum, int firstDataRowNumline, InstrBase instrIf, List<InstrBase> listInstrThen)
    {
        ExecResult execResult = new ExecResult();

        // possible to create the instr?
        if (_coreData.Stage != CoreStage.Build)
        {
            execResult.AddError(new CoreError(ErrorCode.UnableCreateInstrNotInStageBuild, null));
            return execResult;
        }

        if (string.IsNullOrWhiteSpace(excelFileObjectName))
        {
            execResult.AddError(new CoreError(ErrorCode.ExcelFileObjectNameIsNull, null));
            return execResult;
        }

        if (sheetNum < 0)
        {
            execResult.AddError(new CoreError(ErrorCode.SheetNumValueWrong, null));
            return execResult;
        }

        if (firstDataRowNumline < 0)
        {
            execResult.AddError(new CoreError(ErrorCode.FirstDatarowNumLineValueWrong, null));
            return execResult;
        }

        // check the syntax of the excel file object name
        if (!SyntaxUtils.CheckIdSyntax(excelFileObjectName))
        {
            execResult.AddError(new CoreError(ErrorCode.ExcelFileObjectNameSyntaxWrong, null));
            return execResult;
        }

        // check the excelFileObject name, should be defined previsouly
        if (!FindExcelFileObjectName(excelFileObjectName))
        {
            execResult.AddError(new CoreError(ErrorCode.ExcelFileObjectNameDoesNotExists, null));
            return execResult;
        }

        // check instr Then, only SetCellVal is authorized
        if (listInstrThen.Count == 0) 
        {
            execResult.AddError(new CoreError(ErrorCode.AtLeastOneInstrThenExpected, null));
            return execResult;
        }

        InstrForEachRowIfThen instrForIfThen = null;

        // check the instr If, should be allowed
        if(instrIf.InstrType!= InstrType.CompCellVal && instrIf.InstrType != InstrType.CompCellValIsNull)
            execResult.AddError(new CoreError(ErrorCode.IfConditionInstrNotAllowed, instrIf.ToString()));

        // check Then instructions
        foreach (var instrThen in listInstrThen)
        {
            if (instrThen.InstrType != InstrType.SetCellVal && instrThen.InstrType != InstrType.SetCellNull && instrThen.InstrType != InstrType.SetCellBlank)
            {
                execResult.AddError(new CoreError(ErrorCode.ThenConditionInstrNotAllowed, instrThen.ToString()));
            }
        }

        // not allowed instr if or Then?
        if (!execResult.Result)
            return execResult;

        // by default data table header is on the row 0, first data row is 1
        instrForIfThen = new InstrForEachRowIfThen(excelFileObjectName, sheetNum, firstDataRowNumline, instrIf, listInstrThen);
        _coreData.ListInstr.Add(instrForIfThen);
        return execResult;
    }

    /// <summary>
    /// Check the excel Filename, should be defined previsouly
    /// </summary>
    /// <param name="excelFileName"></param>
    /// <returns></returns>
    bool FindExcelFileName(string excelFileName)
    {
        // starting from the last instr, back to the first one
        for (int i = _coreData.ListInstr.Count() - 1; i >= 0; i--)
        {
            // is it an openExcel instr?
            InstrBase instr = _coreData.ListInstr[i];
            InstrOpenExcel instrOpenExcel = instr as InstrOpenExcel;
            if (instrOpenExcel != null)
            {
                if (instrOpenExcel.FileName.Equals(excelFileName, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }
        }

        return false;
    }
    /// <summary>
    /// Check the excelFileObject name, should be defined previsouly
    /// </summary>
    /// <param name="excelFileObjectName"></param>
    /// <returns></returns>
    bool FindExcelFileObjectName(string excelFileObjectName)
    {
        // starting from the last instr, back to the first one
        for(int i = _coreData.ListInstr.Count() - 1;i>=0; i--)
        {
            // is it an openExcel instr?
            InstrBase instr= _coreData.ListInstr[i];
            InstrOpenExcel instrOpenExcel= instr as InstrOpenExcel;
            if (instrOpenExcel != null)
            {
                if (instrOpenExcel.ExcelFileObjectName.Equals(excelFileObjectName, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }
        }

        return false;
    }

}
