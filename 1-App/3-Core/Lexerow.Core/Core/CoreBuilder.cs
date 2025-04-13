using Lexerow.Core.System;
using Lexerow.Core.Utils;
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
    CoreData _coreData;

    public CoreBuilder(CoreData coreData)
    {
        _coreData = coreData;
    }

    public Action<AppTrace> AppTraceEvent;

    /// <summary>
    /// Create an instruction, Cell Value comparison.
    /// exp: CellVal > 12
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompColCellVal CreateInstrCompCellVal(int colNum, ValCompOperator oper, int value)
    {
        // convert the value to an object Value
        ValueInt valueInt = new ValueInt(value);

        InstrCompColCellVal exprComp = new InstrCompColCellVal(colNum, oper, valueInt);

        return exprComp;
    }

    /// <summary>
    /// Create an instruction, Cell Value comparison.
    /// exp: CellVal > 12,3
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompColCellVal CreateInstrCompCellVal(int colNum, ValCompOperator oper, double value)
    {
        // convert the value to an object Value
        ValueDouble valueDouble = new ValueDouble(value);

        InstrCompColCellVal exprComp = new InstrCompColCellVal(colNum, oper, valueDouble);

        return exprComp;
    }

    /// <summary>
    /// Create a Cell Value expression comparison.
    /// exp: CellVal = "georges" 
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompColCellVal CreateInstrCompCellVal(int colNum, ValCompOperator oper, string value)
    {
        // convert the value to an object Value
        ValueString valueString = new ValueString(value);

        InstrCompColCellVal exprComp = new InstrCompColCellVal(colNum, oper, valueString);

        return exprComp;
    }

    /// <summary>
    /// Create an instruction, Cell Value comparison.
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompColCellVal CreateInstrCompCellVal(int colNum, ValCompOperator oper, DateOnly value)
    {
        // convert the value to an object Value
        ValueDateOnly valueDateOnly = new ValueDateOnly(value);

        InstrCompColCellVal exprComp = new InstrCompColCellVal(colNum, oper, valueDateOnly);

        return exprComp;
    }

    /// <summary>
    /// Create an instruction, Cell Value comparison.
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompColCellVal CreateInstrCompCellVal(int colNum, ValCompOperator oper, DateTime value)
    {
        // convert the value to an object Value
        ValueDateTime valueDateTime = new ValueDateTime(value);

        InstrCompColCellVal exprComp = new InstrCompColCellVal(colNum, oper, valueDateTime);

        return exprComp;
    }

    /// <summary>
    /// Create an instruction, Cell Value comparison.
    /// </summary>
    /// <param name="oper"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public InstrCompColCellVal CreateInstrCompCellVal(int colNum, ValCompOperator oper, TimeOnly value)
    {
        // convert the value to an object Value
        ValueTimeOnly valueTimeOnly = new ValueTimeOnly(value);

        InstrCompColCellVal exprComp = new InstrCompColCellVal(colNum, oper, valueTimeOnly);
        return exprComp;
    }

    /// <summary>
    /// Create Comparison instr: B.Cell=Null.
    /// </summary>
    /// <param name="colNum"></param>
    /// <param name="oper"></param>
    /// <returns></returns>
    public InstrCompColCellValIsNull CreateInstrCompCellValIsNull(int colNum)
    {
        return new InstrCompColCellValIsNull(colNum, ValCompOperator.Equal);
    }

    /// <summary>
    /// Create Comparison instr: B.Cell != Null.
    /// </summary>
    /// <param name="colNum"></param>
    /// <param name="oper"></param>
    /// <returns></returns>
    public InstrCompColCellValIsNull CreateInstrCompCellValIsNotNull(int colNum)
    {
        return new InstrCompColCellValIsNull(colNum, ValCompOperator.NotEqual);
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
        SendAppTrace(AppTraceLevel.Info, "CreateInstrOpenExcel: Start");

        // possible to create the instr?
        if (_coreData.Stage != CoreStage.Build)
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

        SendAppTrace(AppTraceLevel.Info, "CreateInstrOpenExcel: End");
        _coreData.ListInstr.Add(instrOpenExcel);
        return execResult;
    }

    /// <summary>
    /// Create instr: If -instrCompIf and... - Then -instrThen-
    /// used in a ForEach Row instruction.
    /// </summary>
    /// <param name="listInstrCompIf"></param>
    /// <param name="instrThen"></param>
    /// <param name="instrIfColThen"></param>
    /// <returns></returns>
    public ExecResult CreateInstrIfColAndThen(List<InstrRetBoolBase> listInstrCompIf, InstrBase instrThen, out InstrIfColThen instrIfColThen)
    {
        List<InstrBase> listInstrThen = new List<InstrBase>() { instrThen };

        return CreateInstrIfColAndThen(listInstrCompIf, listInstrThen, out instrIfColThen);
    }

    /// <summary>
    /// Create instr: If -instrCompIf and... - Then -listOf instrThen-
    /// used in a ForEach Row instruction.
    /// </summary>
    /// <param name="listInstrCompIf"></param>
    /// <param name="listInstrThen"></param>
    /// <param name="instrIfColThen"></param>
    /// <returns></returns>
    public ExecResult CreateInstrIfColAndThen(List<InstrRetBoolBase> listInstrCompIf, List<InstrBase> listInstrThen, out InstrIfColThen instrIfColThen)
    {
        SendAppTrace(AppTraceLevel.Info, "CreateInstrIfColAndThen: Start");
        InstrCompListColCellAnd instrCompListColCellAnd = new InstrCompListColCellAnd(listInstrCompIf);

        ExecResult execResult = new ExecResult();
        instrIfColThen = null;

        //--check instr Then, should be allowed
        foreach (var instrThen in listInstrThen)
        {
            // chkec the instr Then
            if (instrThen.InstrType != InstrType.SetCellVal && instrThen.InstrType != InstrType.SetCellNull && instrThen.InstrType != InstrType.SetCellBlank)
            {
                execResult.AddError(new CoreError(ErrorCode.ThenConditionInstrNotAllowed, instrThen.ToString()));
            }
        }

        // not allowed instr if or Then?
        if (!execResult.Result)
            return execResult;

        // ok, create the IfCol Then instr
        instrIfColThen = new InstrIfColThen();
        instrIfColThen.InstrIf = instrCompListColCellAnd;
        instrIfColThen.ListInstrThen.AddRange(listInstrThen);
        SendAppTrace(AppTraceLevel.Info, "CreateInstrIfColAndThen: End");
        return execResult;
    }

    /// <summary>
    /// Create instr: If -instrCompIf- Then -instrThen-
    /// used in a ForEach Row instruction.
    /// </summary>
    /// <param name="instrCompIf"></param>
    /// <param name="instrThen"></param>
    /// <param name="instrIfColThen"></param>
    /// <returns></returns>
    public ExecResult CreateInstrIfColThen(InstrRetBoolBase instrCompIf, InstrBase instrThen, out InstrIfColThen instrIfColThen)
    {
        List<InstrBase> listInstrThen = new List<InstrBase>
        {
            instrThen
        };
        return CreateInstrIfColThen(instrCompIf, listInstrThen, out instrIfColThen);
    }

    /// <summary>
    /// Create instr: If -instrIf- Then -listOf instrThen-
    /// used in a ForEach Row instruction.
    /// </summary>
    /// <param name="instrCompIf"></param>
    /// <param name="instrSetValThen"></param>
    /// <param name="instrIfColThen"></param>
    /// <returns></returns>
    public ExecResult CreateInstrIfColThen(InstrRetBoolBase instrCompIf, List<InstrBase> listInstrThen, out InstrIfColThen instrIfColThen)
    {
        ExecResult execResult = new ExecResult();
        instrIfColThen = null;

        //--check instr Then, should be allowed
        foreach (var instrThen in listInstrThen)
        {
            // chkec the instr Then
            if (instrThen.InstrType != InstrType.SetCellVal && instrThen.InstrType != InstrType.SetCellNull && instrThen.InstrType != InstrType.SetCellBlank)
            {
                execResult.AddError(new CoreError(ErrorCode.ThenConditionInstrNotAllowed, instrThen.ToString()));
            }
        }

        // not allowed instr if or Then?
        if (!execResult.Result)
            return execResult;

        // ok, create the IfCol Then instr
        instrIfColThen = new InstrIfColThen();
        instrIfColThen.InstrIf = instrCompIf;
        instrIfColThen.ListInstrThen.AddRange(listInstrThen);
        return execResult;
    }

    /// <summary>
    /// Create an instr and save it:
    /// OnExcel
    ///     ForEach Row
    ///         If -instrIf- Then -listOf instrThen-
    /// </summary>
    /// <param name="excelFileObjectName"></param>
    /// <param name="sheetNum"></param>
    /// <param name="firstDataRowNumline"></param>
    /// <param name="listInstrIfColThen"></param>
    /// <returns></returns>
    public ExecResult CreateInstrOnExcelForEachRowIfThen(string excelFileObjectName, int sheetNum, int firstDataRowNumline, InstrIfColThen instrIfColThen)
    {
        List<InstrIfColThen> listInstrIfColThen = new List<InstrIfColThen>
        {
            instrIfColThen
        };
        return CreateInstrOnExcelForEachRowIfThen(excelFileObjectName, sheetNum, firstDataRowNumline, listInstrIfColThen);
    }

    /// <summary>
    /// Create an instr and save it:
    /// ForEach Row
    ///     listOf If -instrIf- Then -listOf instrThen-
    /// </summary>
    /// <param name="excelFileObjectName"></param>
    /// <param name="sheetNum"></param>
    /// <param name="firstDataRowNumline"></param>
    /// <param name="listInstrIfColThen"></param>
    /// <returns></returns>
    public ExecResult CreateInstrOnExcelForEachRowIfThen(string excelFileObjectName, int sheetNum, int firstDataRowNumline, List<InstrIfColThen> listInstrIfColThen)
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
        if (listInstrIfColThen.Count == 0)
        {
            execResult.AddError(new CoreError(ErrorCode.AtLeastOneInstrIfColThenExpected, null));
            return execResult;
        }

        // by default data table header is on the row 0, first data row is 1
        InstrOnExcelForEachRowIfThen instrForEachRowIfThen = new InstrOnExcelForEachRowIfThen(excelFileObjectName, sheetNum, firstDataRowNumline, listInstrIfColThen);
        _coreData.ListInstr.Add(instrForEachRowIfThen);
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


    public void SendAppTrace(AppTraceLevel level, string msg)
    {
        if (AppTraceEvent == null) return;

        AppTrace appTrace = new AppTrace(AppTraceModule.Builder, level, msg);
        AppTraceEvent(appTrace);
    }

}
