using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.Utils;

namespace Lexerow.Core.ProgExec;

/// <summary>
/// Instruction OnExcel Executor.
/// </summary>
internal class InstrOnExcelExecutor
{
    private IActivityLogger _logger;

    private IExcelProcessor _excelProcessor;

    public InstrOnExcelExecutor(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
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
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="listInstr"></param>
    /// <param name="listVar"></param>
    /// <param name="instrOnExcel"></param>
    /// <returns></returns>
    public bool ExecInstrOnExcel(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrOnExcel instrOnExcel)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrOnExcelExecutor.ExecInstrOnExcel", string.Empty);
        bool res;

        // check the init, several cases to manage
        if (!CheckInitOnExcel(result, ctx, progRunVarMgr, instrOnExcel, out bool exitStack))
            return false;

        if (exitStack) return true;

        // save and close the current excel file
        if (ctx.ExcelFileObject != null)
        {
            if (!CloseFileExecutor.Exec(result, _excelProcessor, ctx.ExcelFileObject.ExcelFile))
                return false;
        }

        // next file to process now
        ctx.FileToProcessNum++;

        //no more file to process
        if (ctx.FileToProcessNum >= ctx.ListSelectedFilename.Count)
        {
            // remove the instr OnExcel from the stack
            ctx.StackInstr.Pop();
            return true;
        }
        // run the instr OnExcel on the selected excel file
        var selectedFilename = ctx.ListSelectedFilename[ctx.FileToProcessNum];
        ctx.ExcelFileObject = new InstrExcelFileObject(selectedFilename.InstrBase.FirstScriptToken(), selectedFilename.Filename);

        // update insights
        result.Insights.StartNewFile(selectedFilename.Filename);

        // load the excel file
        if (!OpenExcelFile(result, ctx.ExcelFileObject))
            return false;

        // process all sheets, one by one
        InstrProcessSheets instrProcessSheets = new InstrProcessSheets(instrOnExcel.FirstScriptToken(), instrOnExcel.ListSheets);
        ctx.StackInstr.Push(instrProcessSheets);
        return true;
    }


    private bool CheckInitOnExcel(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrOnExcel instrOnExcel, out bool exitStack)
    {
        exitStack = false;

        if (ctx.FileToProcessNum > -1)
            // not the init stage, continue
            return true;

        //--init-0, case OnExcel "data.xslx"
        if (!IsOnExcelInitFilenameString(result, ctx, instrOnExcel, out exitStack))
            return false;
        if (exitStack) return true;

        //--init-0, case OnExcel files, files can be a string, varname or a selectFiles fctcall
        if (!IsOnExcelInitFilenameVar(result, ctx, progRunVarMgr, instrOnExcel, out exitStack))
            return false;
        if (exitStack) return true;

        //--init-0, case OnExcel "data.xslx", +"file.xlsx", ..
        // instrOnExcel.RunInstrSelectFiles -> stack.Push() -> ctx.PrevInstr=InstrSelectFiles

        //--init-0, all cases checked
        if (instrOnExcel.InstrSelectFiles == null)
        {
            result.AddError(ErrorCode.ExecInstrNotManaged, instrOnExcel.FirstScriptToken());
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
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="instrOnExcel"></param>
    /// <param name="exitStack"></param>
    /// <returns></returns>
    private bool IsOnExcelInitFilenameString(Result result, ProgExecContext ctx, InstrOnExcel instrOnExcel, out bool exitStack)
    {
        exitStack = false;
        if (ctx.PrevInstrExecuted != null) return true;
        if (instrOnExcel.InstrFiles == null) return true;

        var instrValue = instrOnExcel.InstrFiles as InstrValue;
        if (instrValue == null) return true;

        if (instrValue.ValueBase.ValueType != System.ValueType.String)
        {
            // value should be string, a filename
            result.AddError(ErrorCode.ExecInstrVarTypeNotExpected, instrValue.FirstScriptToken());
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
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="progRunVarMgr"></param>
    /// <param name="instrOnExcel"></param>
    /// <param name="exitStack"></param>
    /// <returns></returns>
    private bool IsOnExcelInitFilenameVar(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrOnExcel instrOnExcel, out bool exitStack)
    {
        exitStack = false;
        if (ctx.PrevInstrExecuted != null) return true;
        if (instrOnExcel.InstrFiles == null) return true;

        var instrObjectName = instrOnExcel.InstrFiles as InstrObjectName;
        if (instrObjectName == null) return true;

        // get the final value of the var
        ProgExecVar progRunVar = progRunVarMgr.FindLastInnerVarByName(instrObjectName.ObjectName);
        if (progRunVar == null)
        {
            // var name not found, not defined before in the script
            result.AddError(ErrorCode.ExecInstrVarNotFound, instrOnExcel.InstrFiles.FirstScriptToken());
            return false;
        }

        //--1/ var value is a ConstValue string?
        InstrValue instrValue = progRunVar.Value as InstrValue;
        if (instrValue != null)
        {
            if (instrValue.ValueBase.ValueType != System.ValueType.String)
            {
                // value should be string, a filename
                result.AddError(ErrorCode.ExecInstrVarTypeNotExpected, instrValue.FirstScriptToken());
                return false;
            }

            // create an adhoc SelectFiles instr
            instrOnExcel.InstrSelectFiles = new InstrSelectFiles(instrValue.FirstScriptToken());
            instrOnExcel.InstrSelectFiles.AddParamSelect(instrValue);
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
    /// <param name="result"></param>
    /// <param name="instrValue"></param>
    /// <param name="instrExcelFileObject"></param>
    /// <returns></returns>
    private bool OpenExcelFile(Result result, InstrExcelFileObject instrExcelFileObject)
    {
        // execute the instr OpenExcel(fileName)
        if (!_excelProcessor.Open(instrExcelFileObject.Filename, out IExcelFile excelFile, out ResultError error))
        {
            result.AddError(error);
            return false;
        }
        instrExcelFileObject.ExcelFile = excelFile;
        return true;
    }
}