using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.Func;
using Lexerow.Core.Utils;

namespace Lexerow.Core.InstrProgExec;

/// <summary>
/// Instr SelectFiles runner.
/// Select files regarding parameters.
/// exp:
/// 1/ basic case: SelectFiles("data.xlsx")  select one file.
/// 2/ with joker: SelectFiles("*.xlsx")  select several files.
/// 3/ with var: SelectFiles(filename)
/// </summary>
public class InstrSelectFilesExecutor
{
    private IActivityLogger _logger;

    public InstrSelectFilesExecutor(IActivityLogger activityLogger)
    {
        _logger = activityLogger;
    }

    /// <summary>
    /// Execute instr SelectExcel.
    /// Format:
    ///   SelectExcel("file.xlsx")
    ///   SelectExcel(fileName)     filename="file.xlsx"
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="listInstr"></param>
    /// <param name="listVar"></param>
    /// <param name="instrSelectFiles"></param>
    /// <returns></returns>
    public bool Exec(Result result, ProgExecContext ctx, Program program, InstrFuncSelectFiles instrSelectFiles)
    {
        _logger.LogExecStart(ActivityLogLevel.Info, "InstrSelectFilesRunner.Run", string.Empty);

        // a param (fct call or string concatenation) was processed before
        if (instrSelectFiles.CurrParamNum > -1 && ctx.PrevInstrExecuted != null)
        {
            // save the decoded param
            instrSelectFiles.RunTmpListFinalInstrParams.Add(ctx.PrevInstrExecuted);
            ctx.PrevInstrExecuted = null;
        }

        //--manage->save const value (string) and varname, if a fctcall or string concatenation is found, push it on the stack and process it
        while (true)
        {
            instrSelectFiles.CurrParamNum++;

            // no more param to process
            if (instrSelectFiles.CurrParamNum >= instrSelectFiles.ListInstrParams.Count)
                // next stage: select files, based on selected and unselected files
                break;

            // get the current param and selector
            InstrBase param = instrSelectFiles.ListInstrParams[instrSelectFiles.CurrParamNum];
            InstrFuncSelectFilesSelector selector = instrSelectFiles.ListFilesSelectors[instrSelectFiles.CurrParamNum];

            // is it a const value (string) or varname?
            if (param.InstrType == InstrType.Value || param.InstrType == InstrType.NameObject)
            {
                instrSelectFiles.RunTmpListFinalInstrParams.Add(param);
                continue;
            }

            // fct call, string concatenation, the param have to be processed
            ctx.StackInstr.Push(param);
            return true;
        }

        //--list of params have been all decoded, contains only constValue or varname
        for (int i = 0; i < instrSelectFiles.RunTmpListFinalInstrParams.Count; i++)
        {
            // get the current param and selector
            InstrBase param = instrSelectFiles.ListInstrParams[i];
            InstrFuncSelectFilesSelector selector = instrSelectFiles.ListFilesSelectors[i];

            // the param is a const value or a varname, get final list of filename to process: apply select and unselect filters
            if (!DecodeParam(result, program, instrSelectFiles, param, selector))
                return false;
        }

        // save the instr and remove it from the stack
        ctx.PrevInstrExecuted = instrSelectFiles;
        ctx.StackInstr.Pop();
        return true;
    }

    /// <summary>
    /// Decopde the param, should contains a filename.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="program"></param>
    /// <param name="instrSelectFiles"></param>
    /// <param name="param"></param>
    /// <param name="selector"></param>
    /// <returns></returns>
    private bool DecodeParam(Result result, Program program, InstrFuncSelectFiles instrSelectFiles, InstrBase param, InstrFuncSelectFilesSelector selector)
    {
        // get the filename
        if (!InstrUtils.GetStringFromInstr(result, false, program, param, out bool isValueOrVar, out string value))
            return false;

        // not a value or var, can be: a fct call, a math expr or bool expr
        if (!isValueOrVar) return true;

        if (!SelectFiles(result, instrSelectFiles, StringUtils.RemoveStartEndDoubleQuote(value), out List<string> listFilename))
            return false;

        // no file selected
        if (listFilename.Count == 0)
        {
            result.AddWarning(ErrorCode.ExecNoFileSelected, instrSelectFiles.FirstScriptToken());
            // just a warning but continue
            return true;
        }

        // save list of files
        foreach (string filename in listFilename)
            instrSelectFiles.AddFinalFilename(param, filename);

        return true;
    }

    private bool SelectFiles(Result result, InstrBase instrBase, string filename, out List<string> listFilenameOut)
    {
        listFilenameOut = new List<string>();
        try
        {
            if (string.IsNullOrEmpty(filename))
            {
                result.AddError(ErrorCode.ExecInstrFilenameWrong, instrBase.FirstScriptToken());
                return false;
            }

            string filepath = Path.GetDirectoryName(filename);
            string fullpath = Path.GetFullPath(filename);
            string files = Path.GetFileName(filename);

            if (string.IsNullOrEmpty(filepath))
                filepath = ".";

            if (!Path.Exists(filepath))
            {
                result.AddError(ErrorCode.ExecInstrFilePathWrong, instrBase.FirstScriptToken());
                return false;
            }

            DirectoryInfo d = new DirectoryInfo(filepath);
            foreach (var file in d.GetFiles(files))
            {
                // save the filename, with the full path
                listFilenameOut.Add(file.FullName);
            }

            return true;
        }
        catch (Exception ex)
        {
            result.AddError(ErrorCode.ExecInstrAccessFileWrong, instrBase.FirstScriptToken(), ex);
            return false;
        }
    }
}