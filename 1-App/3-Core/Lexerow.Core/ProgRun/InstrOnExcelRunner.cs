using Lexerow.Core.Core.Exec;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.ProgRun;
using NPOI.HPSF;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgRun;
internal class InstrOnExcelRunner
{
    IActivityLogger _logger;

    IExcelProcessor _excelProcessor;

    public InstrOnExcelRunner(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
    }

    /// <summary>
    /// OnExcel "data.xslx"    
    /// OnExcel file
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="listInstr"></param>
    /// <param name="listVar"></param>
    /// <param name="instrOnExcel"></param>
    /// <returns></returns>
    public bool RunInstrOnExcel(ExecResult execResult, ProgramRunnerContext ctx, List<ExecVar> listVar, InstrOnExcel instrOnExcel)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "ProgRunner.RunInstrOnExcel", string.Empty);
        bool res;

        // define list of excel files to process, one time
        if (ctx.FileToProcessNum == -1)
        {
            // manage excel files, name are: string or varName
            res = GetExcelFilesFromNames(execResult, listVar, instrOnExcel.ListFiles, instrOnExcel, out List<InstrExcelFileObject> listExcelFileNames);
            ctx.ListExcelFileObject.AddRange(listExcelFileNames);
            if (!res) return false;
        }

        // save and close the current excel file
        if (ctx.ExcelFileObject != null) 
        {
            if(!CloseExcelFileRunner.Exec(_excelProcessor, ctx.ExcelFileObject.ExcelFile, out var error))
            {
                execResult.AddError(error);
                return false;
            }
        }

        // next file to process now
        ctx.FileToProcessNum++;

        //no more file to process
        if(ctx.FileToProcessNum >= ctx.ListExcelFileObject.Count)
        {
            // remove the instr OnExcel from the stack
            ctx.StackInstr.Pop();
            return true;
        }
        // run the instr OnExcel on the selected excel file
        var excelFileObject = ctx.ListExcelFileObject[ctx.FileToProcessNum];

        // need to load the excel file object?
        if(excelFileObject.ExcelFile==null)
        {
            // load
            if(!OpenExcelFile(execResult, excelFileObject))
                return false;
        }

        ctx.ExcelFileObject= excelFileObject;

        // process sheets
        InstrNextSheet instrNextSheet = new InstrNextSheet(instrOnExcel.FirstScriptToken(), instrOnExcel.ListSheets);
        ctx.StackInstr.Push(instrNextSheet);
        return true;
    }

    /// <summary>
    /// after OnExcel, comes here, manage sheets to process.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrNextSheet"></param>
    /// <returns></returns>
    public bool RunInstrNextSheet(ExecResult execResult, ProgramRunnerContext ctx, List<ExecVar> listVar, InstrNextSheet instrNextSheet)
    {
        // move to next sheet num
        instrNextSheet.SheetNum++;

        if(instrNextSheet.SheetNum>= instrNextSheet.ListSheet.Count)
        {
            // no more sheet to process, got back to the instr OnExcel
            ctx.StackInstr.Pop();
            return true;
        }

        // focus on the current sheet
        InstrOnSheet instrOnSheet = instrNextSheet.ListSheet[instrNextSheet.SheetNum];
        ctx.StackInstr.Push(instrOnSheet);
        ctx.ExcelSheet = null;
        return true;
    }

    /// <summary>
    /// Process a sheet, execute all defined instr in ForEach Row block.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrOnSheet"></param>
    /// <returns></returns>
    public bool RunInstrOnSheet(ExecResult execResult, ProgramRunnerContext ctx, List<ExecVar> listVar, InstrOnSheet instrOnSheet)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "ProgRunner.RunInstrOnSheet", string.Empty);
        bool res;

        // start of the sheet processing?
        if (ctx.ExcelSheet == null)
        {
            // get the sheet from excel
            ctx.ExcelSheet = _excelProcessor.GetSheetAt(ctx.ExcelFileObject.ExcelFile, instrOnSheet.SheetNum-1);
            
            InstrNextRow instrNextRow = new InstrNextRow(instrOnSheet.FirstScriptToken(), instrOnSheet.ListInstrForEachRow);
            // translate in base0 from human readable base1
            instrNextRow.RowNum= instrOnSheet.FirstRowNum-1;
            instrNextRow.ColNum= instrOnSheet.FirstColNum-1;

            // set to the first datarow
            ctx.StackInstr.Push(instrNextRow);
            return true;
        }

        // go back to the instr Next Sheet
        ctx.StackInstr.Pop();
        return true;
    }

    /// <summary>
    /// proceed next datarow.
    ///  -Stack: NextRow, OnSheet, OnExcel
    ///  Next: ForEachRow.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrNewRow"></param>
    /// <returns></returns>
    public bool RunInstrNextRow(ExecResult execResult, ProgramRunnerContext ctx, List<ExecVar> listVar, InstrNextRow instrNewRow)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "ProgRunner.RunInstrOnSheet", string.Empty);

        // next data row exists?
        int lastRowNum= _excelProcessor.GetLastRowNum(ctx.ExcelSheet);
        if(instrNewRow.RowNum > lastRowNum)
        {
            // no more datarow to process, go back to OnSheet instr
            ctx.StackInstr.Pop();
            return true;

        }

        ctx.RowNum = instrNewRow.RowNum;
        // prepare the next one
        instrNewRow.RowNum++;

        InstrForEachRow instrForEachRow = new InstrForEachRow(instrNewRow.FirstScriptToken(), instrNewRow.ListInstrForEachRow);

        ctx.StackInstr.Push(instrForEachRow);
        return true;
    }


    /// <summary>
    /// Execute next instr defined in the ForEach/Next instr block.
    ///  -Stack: ForEachRow, NextRow, OnSheet, OnExcel
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="listVar"></param>
    /// <param name="instrForEachRow"></param>
    /// <returns></returns>
    public bool RunInstrForEachRow(ExecResult execResult, ProgramRunnerContext ctx, List<ExecVar> listVar, InstrForEachRow instrForEachRow)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "ProgRunner.RunInstrForEachRow", string.Empty);

        // TODO: execute next instr in ForEach Row
        instrForEachRow.InstrToProcessNum++;

        // next instr exists?
        if(instrForEachRow.InstrToProcessNum >= instrForEachRow.ListInstr.Count)
        {
            // no more instr to execute in OnSheet/ForEachRow
            ctx.StackInstr.Pop();
            return true;
        }

        // get the next instr to execute
        InstrBase instrBase = instrForEachRow.ListInstr[instrForEachRow.InstrToProcessNum];
        ctx.StackInstr.Push(instrBase);
        return true;
    }

    /// <summary>
    /// cases:
    ///   OnExcel "data.xslx"    
    ///   OnExcel file
    ///
    /// Not yet implemented:
    ///   OnExcel "*.xslx"
    ///   OnExcel file, "data.xslx", ...
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="listVar"></param>
    /// <param name="listFiles"></param>
    /// <param name="listExcelFiles"></param>
    /// <returns></returns>
    bool GetExcelFilesFromNames(ExecResult execResult, List<ExecVar> listVar, List<InstrBase> listFiles, InstrOnExcel instrOnExcel, out List<InstrExcelFileObject> listExcelFiles)
    {
        listExcelFiles= new List<InstrExcelFileObject>();

        foreach (var item in listFiles)
        {
            //--is it a string, so directly the file name? exp: "data.xlsx"
            InstrConstValue instrConstValue = item as InstrConstValue;
            if (instrConstValue != null)
            {
                // should be a string, a filename
                var value= instrConstValue.ValueBase as ValueString;
                if(value==null)
                {
                    execResult.AddError(ErrorCode.RunInstrTypeStringExpected, instrConstValue.FirstScriptToken());
                    return false;
                }

                InstrExcelFileObject instrExcelFileObject = new InstrExcelFileObject(instrConstValue.FirstScriptToken(), instrConstValue.RawValue);
                listExcelFiles.Add(instrExcelFileObject);
                return true;
            }

            //--is it a varname?  exp: file
            InstrObjectName instrObjectName = item as InstrObjectName;
            if (instrObjectName != null) 
            {
                // get the var, should be an excel file opened before
                var var = listVar.FirstOrDefault(v => v.AreSame(instrObjectName));
                if (var == null)
                {
                    execResult.AddError(ErrorCode.RunInstrVarNotFound, instrObjectName.FirstScriptToken());
                    return false;
                }

                // the var type should be excel file
                InstrExcelFileObject excelFile = var.Value as InstrExcelFileObject;
                if (excelFile == null)
                {
                    execResult.AddError(ErrorCode.RunInstrVarTypeNotExpected, var.Value.FirstScriptToken());
                    return false;
                }
                listExcelFiles.Add(excelFile);
                return true;
            }

            execResult.AddError(ErrorCode.RunInstrNotManaged, item.FirstScriptToken());
            return false;
        }
        execResult.AddError(ErrorCode.RunInstrMissing, instrOnExcel.FirstScriptToken());
        return false;
    }

    /// <summary>
    /// Open the excel file object, from the name.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="instrConstValue"></param>
    /// <param name="instrExcelFileObject"></param>
    /// <returns></returns>
    bool OpenExcelFile(ExecResult execResult, InstrExcelFileObject instrExcelFileObject)
    {
        // execute the instr OpenExcel(fileName)
        if (!_excelProcessor.Open(instrExcelFileObject.Filename, out IExcelFile excelFile, out ExecResultError error))
        {
            execResult.AddError(error);
            return false;
        }
        instrExcelFileObject.ExcelFile = excelFile;
        return true;
    }

}
