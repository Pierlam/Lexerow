using Lexerow.Core.Core.Exec;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.ProgRun;
using NPOI.HPSF;
using NPOI.OpenXmlFormats.Spreadsheet;
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
    /// Process instr OnExcel.
    /// 2 main cases:
    ///   1/ OnExcel "data.xslx"
    ///   2/ OnExcel files
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="listInstr"></param>
    /// <param name="listVar"></param>
    /// <param name="instrOnExcel"></param>
    /// <returns></returns>
    public bool RunInstrOnExcel(ExecResult execResult, ProgramRunnerContext ctx, ProgRunVarMgr progRunVarMgr, InstrOnExcel instrOnExcel)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "ProgRunner.RunInstrOnExcel", string.Empty);
        bool res;

        // check the init, several cases to manage
        if (!CheckInitOnExcel(execResult, ctx, progRunVarMgr, instrOnExcel, out bool exitStack))
            return false;
         
        if(exitStack) return true;

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
        if(ctx.FileToProcessNum >= ctx.ListSelectedFilename.Count)
        {
            // remove the instr OnExcel from the stack
            ctx.StackInstr.Pop();
            return true;
        }
        // run the instr OnExcel on the selected excel file
        var selectedFilename = ctx.ListSelectedFilename[ctx.FileToProcessNum];
        ctx.ExcelFileObject = new InstrExcelFileObject(selectedFilename.InstrBase.FirstScriptToken(), selectedFilename.Filename);

        // load the excel file
        if (!OpenExcelFile(execResult, ctx.ExcelFileObject))
            return false;

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
    public bool RunInstrNextSheet(ExecResult execResult, ProgramRunnerContext ctx, ProgRunVarMgr progRunVarMgr, InstrNextSheet instrNextSheet)
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
    public bool RunInstrOnSheet(ExecResult execResult, ProgramRunnerContext ctx, ProgRunVarMgr progRunVarMgr, InstrOnSheet instrOnSheet)
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
    public bool RunInstrNextRow(ExecResult execResult, ProgramRunnerContext ctx, ProgRunVarMgr progRunVarMgr, InstrNextRow instrNewRow)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "ProgRunner.RunInstrNextRow", string.Empty);

        // next data row exists?
        int lastRowNum= _excelProcessor.GetLastRowNum(ctx.ExcelSheet);
        if(instrNewRow.RowNum > lastRowNum)
        {
            // no more datarow to process, go back to OnSheet instr
            ctx.StackInstr.Pop();
            _logger.LogRunEnd(ActivityLogLevel.Info, "ProgRunner.RunInstrNextRow", "No More row");
            return true;
        }

        ctx.RowNum = instrNewRow.RowNum;
        // prepare the next one
        instrNewRow.RowNum++;

        InstrForEachRow instrForEachRow = new InstrForEachRow(instrNewRow.FirstScriptToken(), instrNewRow.ListInstrForEachRow);

        ctx.StackInstr.Push(instrForEachRow);

        _logger.LogRunEnd(ActivityLogLevel.Info, "ProgRunner.RunInstrNextRow", "NextRowNum: " + instrNewRow.RowNum);
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
    public bool RunInstrForEachRow(ExecResult execResult, ProgramRunnerContext ctx, ProgRunVarMgr progRunVarMgr, InstrForEachRow instrForEachRow)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "ProgRunner.RunInstrForEachRow", string.Empty);

        // execute next instr in ForEach Row
        instrForEachRow.InstrToProcessNum++;

        // next instr exists?
        if(instrForEachRow.InstrToProcessNum >= instrForEachRow.ListInstr.Count)
        {
            // no more instr to execute in OnSheet/ForEachRow
            ctx.StackInstr.Pop();
            _logger.LogRunEnd(ActivityLogLevel.Info, "ProgRunner.RunInstrForEachRow", "No More Instr");
            return true;
        }

        // get the next instr to execute
        InstrBase instrBase = instrForEachRow.ListInstr[instrForEachRow.InstrToProcessNum];
        ctx.StackInstr.Push(instrBase);
        _logger.LogRunEnd(ActivityLogLevel.Info, "ProgRunner.RunInstrForEachRow", "InstrToProcessNum: " + instrForEachRow.InstrToProcessNum);
        return true;
    }

    bool CheckInitOnExcel(ExecResult execResult, ProgramRunnerContext ctx, ProgRunVarMgr progRunVarMgr, InstrOnExcel instrOnExcel, out bool exitStack)
    {
        exitStack = false;

        if (ctx.FileToProcessNum > -1) 
            // not the init stage, continue
            return true;

        //--init-0, case OnExcel "data.xslx"
        if (!IsOnExcelInitFilenameString(execResult, ctx, instrOnExcel, out exitStack))
            return false;
        if (exitStack) return true;

        //--init-0, case OnExcel files, files can be a string, varname or a selectFiles fctcall
        if (!IsOnExcelInitFilenameVar(execResult, ctx, progRunVarMgr, instrOnExcel, out exitStack))
            return false;
        if (exitStack) return true;

        //--init-0, case OnExcel "data.xslx", +"file.xlsx", ..
        // instrOnExcel.RunInstrSelectFiles -> stack.Push() -> ctx.PrevInstr=InstrSelectFiles

        //--init-0, all cases checked
        if(instrOnExcel.InstrSelectFiles==null)
        {
            execResult.AddError(ErrorCode.RunInstrNotManaged, instrOnExcel.FirstScriptToken());
            return false;
        }

        // clear the prev instr executed
        ctx.PrevInstrExecuted = null;

        //--set list of excel files to process
        ctx.ListSelectedFilename.AddRange(instrOnExcel.InstrSelectFiles.ListSelectedFilename);
        return true;
    }


    /// <summary>
    /// Init, case OnExcel "data.xslx"
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="instrOnExcel"></param>
    /// <param name="exitStack"></param>
    /// <returns></returns>
    bool IsOnExcelInitFilenameString(ExecResult execResult, ProgramRunnerContext ctx, InstrOnExcel instrOnExcel, out bool exitStack)
    {
        exitStack = false;
        if (ctx.PrevInstrExecuted != null) return true;
        if (instrOnExcel.InstrFiles == null) return true;

        var instrConstValue = instrOnExcel.InstrFiles as InstrConstValue;
        if(instrConstValue==null) return true;
        
        if (instrConstValue.ValueBase.ValueType != System.ValueType.String) 
        {
            // value should be string, a filename
            execResult.AddError(ErrorCode.RunInstrVarTypeNotExpected, instrConstValue.FirstScriptToken());
            return false;
        }


        // create an adhoc SelectFiles instr
        instrOnExcel.InstrSelectFiles = new InstrSelectFiles(instrOnExcel.InstrFiles.FirstScriptToken());
        instrOnExcel.InstrSelectFiles.AddParamSelect(instrOnExcel.InstrFiles);
        ctx.StackInstr.Push(instrOnExcel.InstrSelectFiles);
        exitStack = true;
        return true;
    }

    /// <summary>
    /// Init, OnExcel files
    /// files varname can be:
    /// 1/ a string, exp: "data.xlsx"
    /// 2/ a selectFiles fct call. exp: files=SelectFiles(...)
    /// 3/ a another var, exp: files= myfiles
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="progRunVarMgr"></param>
    /// <param name="instrOnExcel"></param>
    /// <param name="exitStack"></param>
    /// <returns></returns>
    bool IsOnExcelInitFilenameVar(ExecResult execResult, ProgramRunnerContext ctx, ProgRunVarMgr progRunVarMgr, InstrOnExcel instrOnExcel, out bool exitStack)
    {
        exitStack = false;
        if (ctx.PrevInstrExecuted != null) return true;
        if (instrOnExcel.InstrFiles == null) return true;

        var instrObjectName = instrOnExcel.InstrFiles as InstrObjectName;
        if (instrObjectName == null) return true;

        // get the final value of the var
        ProgRunVar progRunVar = progRunVarMgr.FindLastInnerVarByName(instrObjectName.ObjectName);
        if(progRunVar == null)
        {
            // var name not found, not defined before in the script
            execResult.AddError(ErrorCode.RunInstrVarNotFound, instrOnExcel.InstrFiles.FirstScriptToken());
            return false;
        }

        //--1/ var value is a ConstValue string?
        InstrConstValue instrConstValue= progRunVar.Value as InstrConstValue;
        if (instrConstValue != null) 
        {
            if (instrConstValue.ValueBase.ValueType != System.ValueType.String)
            {
                // value should be string, a filename
                execResult.AddError(ErrorCode.RunInstrVarTypeNotExpected, instrConstValue.FirstScriptToken());
                return false;
            }

            // create an adhoc SelectFiles instr
            instrOnExcel.InstrSelectFiles = new InstrSelectFiles(instrConstValue.FirstScriptToken());
            instrOnExcel.InstrSelectFiles.AddParamSelect(instrConstValue);
            ctx.StackInstr.Push(instrOnExcel.InstrSelectFiles);
            exitStack = true;
            return true;
        }

        //--2/ var value is a SelectFiles fct call ?
        InstrSelectFiles instrSelectFiles = progRunVar.Value as InstrSelectFiles;
        if (instrSelectFiles != null) 
        {
            // Should be already executed before
            instrOnExcel.InstrSelectFiles = instrSelectFiles;
        }

        return true;
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
