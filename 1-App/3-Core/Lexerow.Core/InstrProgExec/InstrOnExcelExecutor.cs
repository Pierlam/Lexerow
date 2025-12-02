using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.Func;
using Lexerow.Core.System.InstrDef.Object;
using Lexerow.Core.System.InstrDef.Process;
using Lexerow.Core.Utils;

namespace Lexerow.Core.InstrProgExec;

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
    public bool ExecInstrOnExcel(Result result, ProgExecContext ctx, Program program, InstrOnExcel instrOnExcel)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrOnExcelExecutor.ExecInstrOnExcel", string.Empty);

        //--starting, init stage
        if (instrOnExcel.ExecStage == InstrOnExcelExecStage.Init)
        {
            if (!ExecInitInstrOnExcel(result, ctx, program, instrOnExcel, out bool exitStack))
                return false;

            // have to execute the instr pushed on the stack
            if (exitStack) return true;
            instrOnExcel.ExecStage = InstrOnExcelExecStage.FilesAreSelected;
        }

        InstrBase prevInstrExecuted = ctx.PrevInstrExecuted;

        // if it's a var, replace the prev by the right value
        // TODO:

        //-- stage filenames are selected
        if (instrOnExcel.ExecStage == InstrOnExcelExecStage.FilesAreSelected)
        {
            // check the previous instr executed
            if(prevInstrExecuted==null)
                return result.AddError(ErrorCode.ExecInstrNotManaged, instrOnExcel.InstrFiles.FirstScriptToken());

            // the prev instr executed type should be FilesSelected
            InstrObjectFilenamesSelected instrObjectFilenamesSelected= prevInstrExecuted as InstrObjectFilenamesSelected;

            if(instrObjectFilenamesSelected==null)
                return result.AddError(ErrorCode.ExecInstrNotManaged, instrOnExcel.InstrFiles.FirstScriptToken());

            // get the list of filename to process
            ctx.ListSelectedFilename.AddRange(instrObjectFilenamesSelected.ListSelectedFilename);
            instrOnExcel.ExecStage = InstrOnExcelExecStage.ProcessFile;
        }

        // save and close the current excel file
        if (ctx.ExcelFileObject != null)
        {
            if (!CloseFileExecutor.Exec(result, _excelProcessor, ctx.ExcelFileObject.ExcelFile))
                return false;
        }

        // next file to process now
        instrOnExcel.FileToProcessNum++;

        //no more file to process
        if (instrOnExcel.FileToProcessNum >= ctx.ListSelectedFilename.Count)
        {
            // remove the instr OnExcel from the stack
            ctx.StackInstr.Pop();
            return true;
        }
        // run the instr OnExcel on the selected excel file
        var selectedFilename = ctx.ListSelectedFilename[instrOnExcel.FileToProcessNum];
        ctx.ExcelFileObject = new InstrObjectExcelFile(selectedFilename.InstrBase.FirstScriptToken(), selectedFilename.Filename);

        // update insights
        result.Insights.StartNewFile(selectedFilename.Filename);

        // load the excel file
        if (!OpenExcelFile(result, ctx.ExcelFileObject))
            return false;

        // process all sheets, one by one
        InstrProcessSheets instrProcessSheets = new InstrProcessSheets(instrOnExcel.FirstScriptToken(), instrOnExcel.ListSheets);
        ctx.StackInstr.Push(instrProcessSheets);
        ctx.PrevInstrExecuted = null;
        return true;
    }

    public bool ExecInstrOnExcel_OLD(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrOnExcel instrOnExcel)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrOnExcelExecutor.ExecInstrOnExcel", string.Empty);
        bool res;

        //// check the init, several cases to manage
        //if (!CheckInitOnExcel(result, ctx, progRunVarMgr, instrOnExcel, out bool exitStack))
        //    return false;

        //if (exitStack) return true;

        // save and close the current excel file
        if (ctx.ExcelFileObject != null)
        {
            if (!CloseFileExecutor.Exec(result, _excelProcessor, ctx.ExcelFileObject.ExcelFile))
                return false;
        }

        // next file to process now
        instrOnExcel.FileToProcessNum++;

        // PB: OnExcel ObjectFilenames, ctx.ListSelectedFilename est vide!! jamais setté
        // revoir ctx.FileToProcessNum  et ctx.ListSelectedFilename   
        // instrOnExcel.InstrObjectFilenamesSelected  contient la liste des noms de fichiers
        //ici();

        //no more file to process
        if (instrOnExcel.FileToProcessNum >= ctx.ListSelectedFilename.Count)
        {
            // remove the instr OnExcel from the stack
            ctx.StackInstr.Pop();
            return true;
        }
        // run the instr OnExcel on the selected excel file
        var selectedFilename = ctx.ListSelectedFilename[instrOnExcel.FileToProcessNum];
        ctx.ExcelFileObject = new InstrObjectExcelFile(selectedFilename.InstrBase.FirstScriptToken(), selectedFilename.Filename);

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


    private bool ExecInitInstrOnExcel(Result result, ProgExecContext ctx, Program program, InstrOnExcel instrOnExcel, out bool exitStack)
    {
        exitStack = false;

        //--the filename to process is a value, exp: OnExcel "dat*.xlsx"
        if(!GetStringFromInstrValue(result, ctx, program, instrOnExcel.InstrFiles, out bool isValue))
            return false;

        // XXX rework code below

        //XXXXXXXXXXX
        if (!InstrUtils.GetStringFromInstrValue(result, false, instrOnExcel.InstrFiles, out bool isValue, out string value))
            return false;

        if (isValue)
        {
            // create an adhoc selectFiles to execute right now
            InstrFuncSelectFiles instrFuncSelectFiles = new InstrFuncSelectFiles(instrOnExcel.InstrFiles.FirstScriptToken());
            instrFuncSelectFiles.AddParamSelect(instrOnExcel.InstrFiles);

            ctx.StackInstr.Push(instrFuncSelectFiles);
            exitStack = true;
            instrOnExcel.ExecStage = InstrOnExcelExecStage.FilesAreSelected;
            return true;
        }

        //--the filename instr to process is a var?
        InstrNameObject instrNameObject = instrOnExcel.InstrFiles as InstrNameObject;
        if (instrNameObject != null) 
        {
            InstrSetVar instrSetVar= program.FindLastVarSet(instrNameObject.Name);
            if(instrSetVar==null)
            {
                result.AddError(ErrorCode.ExecInstrNotManaged, instrOnExcel.InstrFiles.FirstScriptToken());
                return false;
            }
            // get the setVar right instr
            //instrSetVar.InstrRight;
        }

        //--the filename instr to process is a string concat?
        // TODO:

        return result.AddError(ErrorCode.ExecInstrNotManaged, instrOnExcel.InstrFiles.FirstScriptToken());
    }

    private bool CheckInitOnExcel(Result result, ProgExecContext ctx, ProgExecVarMgr progRunVarMgr, InstrOnExcel instrOnExcel, out bool exitStack)
    {
        exitStack = false;

        //if (ctx.FileToProcessNum > -1)
        //    // not the init stage, continue
        //    return true;

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
        if (instrOnExcel.InstrObjectFilenamesSelected == null)
        {
            result.AddError(ErrorCode.ExecInstrNotManaged, instrOnExcel.FirstScriptToken());
            return false;
        }

        // clear the prev instr executed
        ctx.PrevInstrExecuted = null;

        //--set list of excel files to process
        //ctx.ListSelectedFilename.AddRange(instrOnExcel.InstrSelectFiles.ListSelectedFilename);
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
        //instrOnExcel.InstrSelectFiles = new InstrFuncSelectFiles(instrOnExcel.InstrFiles.FirstScriptToken());
        //instrOnExcel.InstrSelectFiles.AddParamSelect(instrOnExcel.InstrFiles);

        InstrFuncSelectFiles instrFuncSelectFiles = new InstrFuncSelectFiles(instrOnExcel.InstrFiles.FirstScriptToken());
        instrFuncSelectFiles.AddParamSelect(instrOnExcel.InstrFiles);

        //ctx.StackInstr.Push(instrOnExcel.InstrSelectFiles);
        ctx.StackInstr.Push(instrFuncSelectFiles);
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

        var instrNameObject = instrOnExcel.InstrFiles as InstrNameObject;
        if (instrNameObject == null) return true;

        // get the final value of the var
        ProgExecVar progRunVar = progRunVarMgr.FindLastInnerVarByName(instrNameObject.Name);
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
            //instrOnExcel.InstrSelectFiles = new InstrFuncSelectFiles(instrValue.FirstScriptToken());
            //instrOnExcel.InstrSelectFiles.AddParamSelect(instrValue);
            InstrFuncSelectFiles instrFuncSelectFiles = new InstrFuncSelectFiles(instrValue.FirstScriptToken());
            instrFuncSelectFiles.AddParamSelect(instrValue);

            //ctx.StackInstr.Push(instrOnExcel.InstrSelectFiles);
            ctx.StackInstr.Push(instrFuncSelectFiles);
            exitStack = true;
            return true;
        }

        //--2/ var value is a SelectFiles fct call ?
        InstrObjectFilenamesSelected instrObjectFilenamesSelected= progRunVar.Value as InstrObjectFilenamesSelected;
        if (instrObjectFilenamesSelected != null)
        {
            // Should be already executed before
            //instrOnExcel.InstrSelectFiles = instrObjectFilenamesSelected;
            instrOnExcel.InstrObjectFilenamesSelected = instrObjectFilenamesSelected;
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
    private bool OpenExcelFile(Result result, InstrObjectExcelFile instrExcelFileObject)
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