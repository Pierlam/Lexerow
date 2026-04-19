using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.FuncCall;
using Lexerow.Core.System.InstrDef.Object;
using Lexerow.Core.System.InstrDef.Process;
using Lexerow.Core.Utils;
using OpenExcelSdk;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Instruction OnExcel Executor.
/// </summary>
internal class InstrOnExcelExecutor
{
    private IActivityLogger _logger;

    private ExcelProcessor _excelProcessor;

    public InstrOnExcelExecutor(IActivityLogger activityLogger, ExcelProcessor excelProcessor)
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
    public bool ExecInstrOnExcel(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrOnExcel instrOnExcel)
    {
        _logger.Log(ActivityLogLevel.Debug, "InstrOnExcelExecutor.ExecInstrOnExcel", string.Empty);

        InstrObjectSelectedFiles instrObjectSelectedFiles = null;

        //--starting, init stage
        if (instrOnExcel.ExecStage == InstrOnExcelExecStage.Init)
        {
            if (!ExecInitInstrOnExcel(result, ctx, progExecVarMgr, instrOnExcel, out bool exitStack, out instrObjectSelectedFiles))
                return false;

            // have to execute the instr pushed on the stack
            if (exitStack) return true;
        }

        //--files have to be selected
        if (instrOnExcel.ExecStage == InstrOnExcelExecStage.FilesToSelect)
        {
            InstrBase prevInstrExecuted = ctx.PrevInstrExecuted;

            // check the previous instr executed
            if (prevInstrExecuted == null)
                return result.AddError(ErrorCode.ExecInstrNotManaged, instrOnExcel.InstrFiles.FirstScriptToken());

            // the prev instr executed type should be FilesSelected
            instrObjectSelectedFiles = prevInstrExecuted as InstrObjectSelectedFiles;

            if (instrObjectSelectedFiles == null)
                return result.AddError(ErrorCode.ExecInstrNotManaged, instrOnExcel.InstrFiles.FirstScriptToken());

            // get the list of filename to process
            ctx.ListObjectSelectedFiles.AddRange(instrObjectSelectedFiles.ListObjectSelectedFile);
            instrOnExcel.ExecStage = InstrOnExcelExecStage.ProcessFile;
        }

        //--files are selected
        if (instrOnExcel.ExecStage == InstrOnExcelExecStage.FilesAreSelected)
        {
            // todo: recup data dans ExecInitInstrOnExcel()
            ctx.ListObjectSelectedFiles.AddRange(instrObjectSelectedFiles.ListObjectSelectedFile);
            instrOnExcel.ExecStage = InstrOnExcelExecStage.ProcessFile;
        }

        // can continue only if Stage is ProcessFile
        if (instrOnExcel.ExecStage != InstrOnExcelExecStage.ProcessFile)
            return result.AddError(ErrorCode.ExecInstrNotManaged, instrOnExcel.InstrFiles.FirstScriptToken());

        // save and close the current excel file
        if (ctx.ExcelFileObject != null)
        {
            if (!CloseFileExecutor.Exec(result, _excelProcessor, ctx.ExcelFileObject.ExcelFile))
                return false;
        }

        // next file to process now
        instrOnExcel.FileToProcessNum++;

        //no more file to process
        if (instrOnExcel.FileToProcessNum >= ctx.ListObjectSelectedFiles.Count)
        {
            // remove the instr OnExcel from the stack
            ctx.StackInstr.Pop();
            return true;
        }
        // run the instr OnExcel on the selected excel file
        var selectedFilename = ctx.ListObjectSelectedFiles[instrOnExcel.FileToProcessNum];
        ctx.ExcelFileObject = new InstrObjectExcelFile(selectedFilename.InstrBase.FirstScriptToken(), selectedFilename.Filename);

        // update insights
        result.Insights.StartNewFile(selectedFilename.Filename);

        _logger.Log(ActivityLogLevel.Debug, "InstrOnExcelExecutor.ExecInstrOnExcel.ProcessFile", selectedFilename.Filename);


        // load the excel file
        if (!OpenExcelFile(result, ctx.ExcelFileObject))
            return false;

        // process all sheets, one by one
        InstrProcessSheets instrProcessSheets = new InstrProcessSheets(instrOnExcel.FirstScriptToken(), instrOnExcel.ListSheets);
        ctx.StackInstr.Push(instrProcessSheets);
        ctx.PrevInstrExecuted = null;
        return true;
    }

    /// <summary>
    /// several cases:
    ///   1/ OnExcel "dat.xslx"
    ///   2/ OnExcel SelectFiles()
    ///   3/ files=SelectFiles(), OnExcel files
    ///   4/ files="dat.xslx", OnExcel files
    ///   5/ files "dat*" + ".xlsx", OnExcel files  stringConcat case
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="program"></param>
    /// <param name="instrOnExcel"></param>
    /// <param name="exitStack"></param>
    /// <returns></returns>
    private bool ExecInitInstrOnExcel(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrOnExcel instrOnExcel, out bool exitStack, out InstrObjectSelectedFiles instrObjectSelectedFiles)
    {
        exitStack = false;
        instrObjectSelectedFiles = null;

        InstrBase instrFiles = instrOnExcel.InstrFiles;

        //--case-2,3,4: is it OnExcel files, a var?
        InstrNameObject instrNameObject = instrOnExcel.InstrFiles as InstrNameObject;
        if (instrNameObject != null)
        {
            ProgExecVar progExecVar = progExecVarMgr.FindVarByName(instrNameObject.Name);
            if (progExecVar == null)
                return result.AddError(ErrorCode.ExecInstrNotManaged, instrOnExcel.InstrFiles.FirstScriptToken());

            // it's a var, so get the value of th evar
            instrFiles = progExecVar.Value;
        }

        //--selected files are defined ?
        instrObjectSelectedFiles = instrFiles as InstrObjectSelectedFiles;
        if (instrObjectSelectedFiles != null)
        {
            instrOnExcel.ExecStage = InstrOnExcelExecStage.FilesAreSelected;
            return true;
        }

        //--case-1: The filename to process is a value, exp: OnExcel "dat*.xlsx"
        if (!ManageFilenameIsValue(result, ctx, instrOnExcel, instrFiles, out exitStack))
            return false;
        if (exitStack) return true;

        //--The files param is the SelectFiles instr?
        InstrFuncCallSelectFiles instrFuncSelectFiles = instrFiles as InstrFuncCallSelectFiles;
        if (instrFuncSelectFiles != null)
        {
            ctx.StackInstr.Push(instrFuncSelectFiles);
            instrOnExcel.ExecStage = InstrOnExcelExecStage.FilesToSelect;
            exitStack = true;
            return true;
        }

        //--the filename instr to process is a string concat?
        // TODO:

        return result.AddError(ErrorCode.ExecInstrNotManaged, instrOnExcel.InstrFiles.FirstScriptToken());
    }

    /// <summary>
    ///  The filename to process is a value, exp: OnExcel "dat*.xlsx"
    ///  Create an adhoc instr to process the raw filename.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="program"></param>
    /// <param name="instrOnExcel"></param>
    /// <param name="isValue"></param>
    /// <returns></returns>
    private bool ManageFilenameIsValue(Result result, ProgExecContext ctx, InstrOnExcel instrOnExcel, InstrBase instrBase, out bool isValue)
    {
        isValue = false;

        // the value is a string? a raw filename
        if (!InstrUtils.GetStringFromInstrValue(result, false, instrBase, out isValue, out string value))
            return false;

        if (!isValue) return true;

        // create an adhoc selectFiles to execute right now
        InstrFuncCallSelectFiles instrFuncSelectFiles = new InstrFuncCallSelectFiles(instrBase.FirstScriptToken());
        instrFuncSelectFiles.AddParamSelect(instrBase);

        ctx.StackInstr.Push(instrFuncSelectFiles);
        instrOnExcel.ExecStage = InstrOnExcelExecStage.FilesToSelect;
        return true;
    }

    /// <summary>
    /// Open the excel file object, from the name.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="instrValue"></param>
    /// <param name="instrExcelFileObject"></param>
    /// <returns></returns>
    private bool OpenExcelFile(Result result, InstrObjectExcelFile instrExcelFileObject)
    {
        // execute the instr OpenExcel(fileName)

        try
        {
            ExcelFile excelFile = _excelProcessor.OpenExcelFile(instrExcelFileObject.Filename);
            instrExcelFileObject.ExcelFile = excelFile;
            return true;
        }
        catch (Exception ex)
        {
            result.AddError(ErrorCode.ExecUnableOpenFile, ex, instrExcelFileObject.Filename);
            return false;
        }
    }
}