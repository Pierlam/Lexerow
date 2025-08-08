using Lexerow.Core.Scripts;
using Lexerow.Core.Scripts.SyntaxAnalyze;
using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using Lexerow.Core.Utils;
using NPOI.HPSF;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core;

/// <summary>
/// Program builder.
/// By adding instructions.
/// don't manage script/source code.
/// </summary>
public class ProgBuilder
{
    /// <summary>
    /// Contains programs
    /// </summary>
    CoreData _coreData;

    LexicalAnalyzer _lexicalAnalyzer;

    public ProgBuilder(CoreData coreData)
    {
        _coreData = coreData;
        _lexicalAnalyzer= new LexicalAnalyzer();
    }

    public Action<AppTrace> AppTraceEvent;

    /// <summary>
    /// Create a new program.
    /// Check the name
    /// Becomes the current program.
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    public ExecResult CreateProgram(string name)
    {
        ExecResult execResult = new ExecResult();

        // check the name syntax
        if (!ItemsCheckUtils.CheckIdSyntax(name))
        {
            execResult.AddError(new ExecResultError(ErrorCode.ProgramWrongName, name));
            return execResult;
        }

        // the name should not be already used
        if (_coreData.GetProgramByName(name) != null) 
        {
            execResult.AddError(new ExecResultError(ErrorCode.ProgramNameAlreadyUsed, name));
            return execResult;
        }

        // create the program and save it
        ProgramInstr programInstr = new ProgramInstr(name);
        _coreData.ListProgram.Add(programInstr);

        // becomes the current one
        _coreData.CurrProgramInstr = programInstr;
        return execResult;
    }

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
    /// Create Comparison instr: A.Cell In [ "y", "ok", "oui" ]
    /// In /I ->Case Insensitive  
    /// A.Cell, listOfItems, In=true/NotIn=false, case sensitive=true
    /// </summary>
    /// <param name="colNum"></param>
    /// <param name="listItems"></param>
    /// <param name="inOrNotIn"></param>
    /// <param name="caseSensitive"></param>
    /// <returns></returns>
    public ExecResult CreateInstrCompCellInItems(int colNum, List<string> listItems, bool caseSensitive, out InstrCompColCellInStringItems instr)
    {
        ExecResult execResult = new ExecResult();

        // list should contains one item at list
        if (listItems == null || listItems.Count == 0)
        {
            instr = null;
            execResult.AddError(new ExecResultError(ErrorCode.EntryListIsEmpty, "InsrCompColCellInItems"));
            return execResult;
        }

        // check the items
        if (!ItemsCheckUtils.CheckItemsUnique(listItems))
        {
            instr = null;
            execResult.AddError(new ExecResultError(ErrorCode.ItemsShouldBeNotNullAndUnique, "InsrCompColCellInItems"));
            return execResult;
        }

        instr = new InstrCompColCellInStringItems(colNum,  true, caseSensitive, listItems);
        return execResult;
    }

    /// <summary>
    /// Create Comparison instr: A.Cell NOT In [ "y", "ok", "oui" ]
    /// In /I ->Case Insensitive  
    /// A.Cell, listOfItems, In=true/NotIn=false, case sensitive=true
    /// </summary>
    /// <param name="colNum"></param>
    /// <param name="listItems"></param>
    /// <param name="caseSensitive"></param>
    /// <returns></returns>
    public ExecResult CreateInstrCompCellNotInItems(int colNum, List<string> listItems, bool caseSensitive, out InstrCompColCellInStringItems instr)
    {
        ExecResult execResult = new ExecResult();

        // list should contains one item at list
        if(listItems==null || listItems.Count==0)
        {
            instr = null;
            execResult.AddError(new ExecResultError(ErrorCode.EntryListIsEmpty, "InsrCompColCellNotInItems"));
            return execResult;
        }

        // check the items
        if (!ItemsCheckUtils.CheckItemsUnique(listItems))
        {
            instr = null;
            execResult.AddError(new ExecResultError(ErrorCode.ItemsShouldBeNotNullAndUnique, "InsrCompColCellNotInItems"));
            return execResult;
        }

        instr = new InstrCompColCellInStringItems(colNum, false, caseSensitive, listItems);
        return execResult;
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
    /// 
    /// file -> excelFileObjectName 
    /// </summary>
    /// <param name="fileObjectRet"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public ExecResult CreateInstrOpenExcelParamConst(string varName, string strfileName)
    {
        ExecResult execResult = new ExecResult();
        SendAppTrace(AppTraceLevel.Info, "CreateInstrOpenExcel: Start");

        // no current program, error
        if(_coreData.CurrProgramInstr == null)
        {
            execResult.AddError(new ExecResultError(ErrorCode.NoCurrentProgramExist));
            return execResult;
        }

        // possible to create the instr?
        if (_coreData.CurrProgramInstr.Stage != CoreStage.Build)
        {
            execResult.AddError(new ExecResultError(ErrorCode.UnableCreateInstrNotInStageBuild, "OpenExcel"));
            return execResult;
        }

        if (string.IsNullOrWhiteSpace(varName))
        {
            execResult.AddError(new ExecResultError(ErrorCode.FileObjectNameNullOrEmpty, "OpenExcel"));
            return execResult;
        }

        if (string.IsNullOrWhiteSpace(strfileName))
        {
            execResult.AddError(new ExecResultError(ErrorCode.FileNameNullOrEmpty));
            return execResult;
        }

        // check the syntax of the excel file object name
        if (!ItemsCheckUtils.CheckIdSyntax(varName)) 
        {
            execResult.AddError(new ExecResultError(ErrorCode.ExcelFileObjectNameSyntaxWrong, varName));
            return execResult;
        }

        // check the excelFileObject name, should not be defined previsouly
        //if (FindExcelFileObjectName(varName))
        //{
        //    execResult.AddError(new ExecResultError(ErrorCode.VarNameNotFound, varName));
        //    return execResult;
        //}

        // check the excel filename, should not be defined previously
        //if (FindExcelFileName(strfileName))
        //{
        //    execResult.AddError(new ExecResultError(ErrorCode.FileNameNotFound, strfileName));
        //    return execResult;
        //}

        // create the instruction SetVar
        InstrSetVar instrSetVar = new InstrSetVar();
        instrSetVar.VarName = varName;
        _coreData.CurrProgramInstr.ListInstr.Add(instrSetVar);

        // create the instruction OpenExcel
        InstrOpenExcel instrOpenExcel = new InstrOpenExcel();
        _coreData.CurrProgramInstr.ListInstr.Add(instrOpenExcel);

        // add OpenExcel function open bracket
        _coreData.CurrProgramInstr.ListInstr.Add(new InstrOpenBracket());

        // add the string filename parameter
        InstrConstValue instrConstValue = new InstrConstValue(strfileName);
        _coreData.CurrProgramInstr.ListInstr.Add(instrConstValue);

        // add the close bracket
        _coreData.CurrProgramInstr.ListInstr.Add(new InstrCloseBracket());

        // add an end of line
        _coreData.CurrProgramInstr.ListInstr.Add(new InstrEol());

        SendAppTrace(AppTraceLevel.Info, "CreateInstrOpenExcel: End");
        return execResult;
    }

    /// <summary>
    /// Create the instruction:
    ///     OnExcel fileNameOrVar OnSheet numSheet, firstRow
    ///     
    /// filenameOrVar can be: 
    ///   -filename, with quotes, exp: "myfile.xslx"
    ///   -var, without quote, exp: file   
    ///     the var should be defined before.
    /// </summary>
    /// <param name="numSheet"></param>
    /// <param name="firstRow"></param>
    /// <param name="buildInstrOnExcel"></param>
    /// <returns></returns>
    public ExecResult CreateInstrOnExcelOnSheet(string fileNameOrVar, int numSheet, int firstRow, out BuildInstrOnExcelOnSheet buildInstrOnExcel)
    {
        ExecResult execResult = new ExecResult();

        // check the filename or var
        if (string.IsNullOrWhiteSpace(fileNameOrVar)) 
        {
            execResult.AddError(new ExecResultError(ErrorCode.FileNameNullOrEmpty));
            buildInstrOnExcel = null;
            return execResult;
        }

        // filename or var?
        if(!StringUtils.HasStartEndQuotes(fileNameOrVar, out bool isString))
        {
            execResult.AddError(new ExecResultError(ErrorCode.FileNameWrongSyntax));
            buildInstrOnExcel = null;
            return execResult;
        }

        // if it's a var, should be defined before
        if(!isString)
        {
            // the var should be defined before

        }

        // check numSheet value
        if(numSheet<-1)
        {
            execResult.AddError(new ExecResultError(ErrorCode.IntMustBePositive,"NumSheet"));
            buildInstrOnExcel = null;
            return execResult;
        }

        // check firstRow value
        if (firstRow < -1)
        {
            execResult.AddError(new ExecResultError(ErrorCode.IntMustBePositive, "FirstRow"));
            buildInstrOnExcel = null;
            return execResult;
        }

        // ok, create the build instruction
        buildInstrOnExcel = new BuildInstrOnExcelOnSheet();
        buildInstrOnExcel.SheetNum = numSheet;
        buildInstrOnExcel.FirstRow= firstRow;
        buildInstrOnExcel.FileName = fileNameOrVar;
        buildInstrOnExcel.IsFileNameVar= !isString;
        return execResult;
    }

    /// <summary>
    /// Build ForEach Row -> If part: If D.Cell>50 
    /// </summary>
    /// <param name="buildInstrOnExcel"></param>
    /// <param name="leftPart"></param>
    /// <param name="comp"></param>
    /// <param name="leftValue"></param>
    /// <param name="buildInstrIf"></param>
    /// <returns></returns>
    public ExecResult CreateInstrIf(BuildInstrOnExcelOnSheet buildInstrOnExcel, string excelColName, string excelFuncName, string compOperator, int rightValue, out BuildInstrOnExcelIf buildInstrIf)
    {
        ExecResult execResult = new ExecResult();
        buildInstrIf = null;

        if (string.IsNullOrEmpty(excelColName)) 
        {
            execResult.AddError(new ExecResultError(ErrorCode.ExcelColNameIsEmpty, "IfExcelColName"));
            return execResult;
        }

        if (string.IsNullOrEmpty(excelFuncName))
        {
            execResult.AddError(new ExecResultError(ErrorCode.ExcelFuncNameIsEmpty, "IfExcelColName"));
            return execResult;
        }

        if (string.IsNullOrEmpty(compOperator))
        {
            execResult.AddError(new ExecResultError(ErrorCode.CompOperatorIsEmpty, "IfCompOperator"));
            return execResult;
        }

        // convert the excel column name
        int excelColValue= ExcelExtendedUtils.ColumnNameToNumber(excelColName);
        if (excelColValue < 0) 
        {
            execResult.AddError(new ExecResultError(ErrorCode.ExcelColNameIsWrong, "excelColName"));
            return execResult;
        }

        // convert the excel function name -> Only Cell allowed!
        excelFuncName = excelFuncName.Trim();
        if (!excelFuncName.Equals("Cell", StringComparison.CurrentCultureIgnoreCase))
        {
            execResult.AddError(new ExecResultError(ErrorCode.ExcelFuncNameIsWrong, excelFuncName));
            return execResult;
        }

        // convert the comparison operator
        if(!CompOperatorDecoder.Do(compOperator, out ExecTokCompOperator execTokCompOperator))
        {
            execResult.AddError(new ExecResultError(ErrorCode.CompOperatorIsWrong, compOperator));
            return execResult;
        }

        InstrConstValue instrConstValue = new InstrConstValue(rightValue);

        // create the build instruction, exp: IF A.Cell > 12
        buildInstrIf = new BuildInstrOnExcelIf(new ExecTokExcelCellAddress(excelColName,excelColValue), new ExecTokExcelFuncCell(), execTokCompOperator, instrConstValue);

        return execResult;
    }

    /// <summary>
    /// Build ForEach Row Then part: Then D.Cell= 12
    /// </summary>
    /// <param name="BuildInstrOnExcel"></param>
    /// <param name="buildInstrIf"></param>
    /// <param name="leftPart"></param>
    /// <param name="comp"></param>
    /// <param name="rightValue"></param>
    /// <returns></returns>
    public ExecResult  CreateInstrThenSetVar(BuildInstrOnExcelOnSheet buildInstrOnExcel, BuildInstrOnExcelIf buildInstrIf, string excelColName, string excelFuncName, int rightValue)
    {
        ExecResult execResult = new ExecResult();

        if (string.IsNullOrEmpty(excelColName))
        {
            execResult.AddError(new ExecResultError(ErrorCode.ExcelColNameIsEmpty, "IfExcelColName"));
            return execResult;
        }

        if (string.IsNullOrEmpty(excelFuncName))
        {
            execResult.AddError(new ExecResultError(ErrorCode.ExcelFuncNameIsEmpty, "IfExcelColName"));
            return execResult;
        }

        // convert the excel column name
        int excelColValue = ExcelExtendedUtils.ColumnNameToNumber(excelColName);
        if (excelColValue < 0)
        {
            execResult.AddError(new ExecResultError(ErrorCode.ExcelColNameIsWrong, "excelColName"));
            return execResult;
        }

        // convert the excel function name -> Only Cell allowed!
        excelFuncName = excelFuncName.Trim();
        if (!excelFuncName.Equals("Cell", StringComparison.CurrentCultureIgnoreCase))
        {
            execResult.AddError(new ExecResultError(ErrorCode.ExcelFuncNameIsWrong, excelFuncName));
            return execResult;
        }

        InstrConstValue instrConstValue = new InstrConstValue(rightValue);


        // create the then part instruction 
        //BuildInstrOnExcelThen
            
        // create the build instruction, exp: IF A.Cell > 12
        //buildInstrIf = new BuildInstrIf(new ExecTokExcelCellAddress(excelColName, excelColValue), new ExecTokExcelFuncCell(), execTokCompOperator, instrConstValue);


        // TODO: InstrSetVar, ExcelCellFunc
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
                execResult.AddError(new ExecResultError(ErrorCode.ThenConditionInstrNotAllowed, instrThen.ToString()));
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
                execResult.AddError(new ExecResultError(ErrorCode.ThenConditionInstrNotAllowed, instrThen.ToString()));
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

        // no current program, error
        if (_coreData.CurrProgramInstr == null)
        {
            execResult.AddError(new ExecResultError(ErrorCode.NoCurrentProgramExist));
            return execResult;
        }

        // possible to create the instr?
        if (_coreData.CurrProgramInstr.Stage != CoreStage.Build)
        {
            execResult.AddError(new ExecResultError(ErrorCode.UnableCreateInstrNotInStageBuild, null));
            return execResult;
        }

        if (string.IsNullOrWhiteSpace(excelFileObjectName))
        {
            execResult.AddError(new ExecResultError(ErrorCode.ExcelFileObjectNameIsNull, null));
            return execResult;
        }

        if (sheetNum < 0)
        {
            execResult.AddError(new ExecResultError(ErrorCode.SheetNumValueWrong, null));
            return execResult;
        }

        if (firstDataRowNumline < 0)
        {
            execResult.AddError(new ExecResultError(ErrorCode.FirstDatarowNumLineValueWrong, null));
            return execResult;
        }

        // check the syntax of the excel file object name
        if (!ItemsCheckUtils.CheckIdSyntax(excelFileObjectName))
        {
            execResult.AddError(new ExecResultError(ErrorCode.ExcelFileObjectNameSyntaxWrong, null));
            return execResult;
        }

        // check the excelFileObject name, should be defined previsouly
        //if (!FindExcelFileObjectName(excelFileObjectName))
        //{
        //    execResult.AddError(new ExecResultError(ErrorCode.ExcelFileObjectNameDoesNotExists, null));
        //    return execResult;
        //}

        // check instr Then, only SetCellVal is authorized
        if (listInstrIfColThen.Count == 0)
        {
            execResult.AddError(new ExecResultError(ErrorCode.AtLeastOneInstrIfColThenExpected, null));
            return execResult;
        }

        // by default data table header is on the row 0, first data row is 1
        InstrOnExcelForEachRowIfThen instrForEachRowIfThen = new InstrOnExcelForEachRowIfThen(excelFileObjectName, sheetNum, firstDataRowNumline, listInstrIfColThen);
        _coreData.CurrProgramInstr.ListInstr.Add(instrForEachRowIfThen);
        return execResult;
    }

    /// <summary>
    /// Check the excel Filename, should be defined previsouly
    /// </summary>
    /// <param name="excelFileName"></param>
    /// <returns></returns>
    //bool FindExcelFileName(string excelFileName)
    //{
    //    // no current program, error
    //    if (_coreData.CurrProgramInstr == null)
    //        return false;

    //    // starting from the last instr, back to the first one
    //    for (int i = _coreData.CurrProgramInstr.ListInstr.Count() - 1; i >= 0; i--)
    //    {
    //        // is it an openExcel instr?
    //        InstrBase instr = _coreData.CurrProgramInstr.ListInstr[i];
    //        InstrOpenExcel instrOpenExcel = instr as InstrOpenExcel;
    //        if (instrOpenExcel != null)
    //        {
    //            if (instrOpenExcel.Param.InstrType != InstrType.Var)
    //                // humm, case pb not managed
    //                return false;

    //            InstrVar instrVar = instrOpenExcel.Param as InstrVar;

    //            if (instrVar.VarName.Equals(excelFileName, StringComparison.InvariantCultureIgnoreCase))
    //                return true;
    //        }
    //    }

    //    return false;
    //}

    /// <summary>
    /// Check the excelFileObject name, should be defined previsouly
    /// </summary>
    /// <param name="excelFileObjectName"></param>
    /// <returns></returns>
    //bool FindExcelFileObjectName(string excelFileObjectName)
    //{
    //    // no current program, error
    //    if (_coreData.CurrProgramInstr == null)
    //        return false;

    //    // starting from the last instr, back to the first one
    //    for (int i = _coreData.CurrProgramInstr.ListInstr.Count() - 1; i >= 0; i--)
    //    {
    //        // is it an openExcel instr?
    //        InstrBase instr = _coreData.CurrProgramInstr.ListInstr[i];
    //        InstrOpenExcel instrOpenExcel = instr as InstrOpenExcel;
    //        if (instrOpenExcel != null)
    //        {
    //            if (instrOpenExcel.Param.InstrType != InstrType.Var)
    //                // humm, case pb not managed
    //                return false;

    //            InstrVar instrVar = instrOpenExcel.Param as InstrVar;

    //            if (instrVar.VarName.Equals(excelFileObjectName, StringComparison.InvariantCultureIgnoreCase))
    //                return true;
    //        }
    //    }

    //    return false;
    //}


    public void SendAppTrace(AppTraceLevel level, string msg)
    {
        if (AppTraceEvent == null) return;

        AppTrace appTrace = new AppTrace(AppTraceModule.Builder, level, msg);
        AppTraceEvent(appTrace);
    }

}
