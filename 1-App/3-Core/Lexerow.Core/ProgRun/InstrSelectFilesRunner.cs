using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.Excel;
using Lexerow.Core.System.ProgRun;
using Lexerow.Core.Utils;
using NPOI.HPSF;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgRun;

/// <summary>
/// Instr SelectFiles runner.
/// Select files regarding parameters.
/// exp: 
/// 1/ basic case: SelectFiles("data.xlsx")  select one file.
/// 2/ with joker: SelectFiles("*.xlsx")  select several files.
/// 3/ with var: SelectFiles(filename)  
/// </summary>
public class InstrSelectFilesRunner
{
    IActivityLogger _logger;

    IExcelProcessor _excelProcessor;

    public InstrSelectFilesRunner(IActivityLogger activityLogger, IExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor = excelProcessor;
    }

    /// <summary>
    /// Execute instr SelectExcel.
    /// Format:
    ///   SelectExcel("file.xlsx")
    ///   SelectExcel(fileName)     filename="file.xlsx"
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="ctx"></param>
    /// <param name="listInstr"></param>
    /// <param name="listVar"></param>
    /// <param name="instrSelectFiles"></param>
    /// <returns></returns>
    public bool Run(ExecResult execResult, ProgramRunnerContext ctx, ProgRunVarMgr progRunVarMgr, InstrSelectFiles instrSelectFiles)
    {
        _logger.LogRunStart(ActivityLogLevel.Info, "InstrSelectFilesRunner.Run", string.Empty);

        // a param (fct call or string concatenation) was processed before
        if(instrSelectFiles.CurrParamNum>-1 && ctx.PrevInstrExecuted!=null)
        {
            // save the decoded param
            instrSelectFiles.RunTmpListFinalInstrParams.Add(ctx.PrevInstrExecuted);
            ctx.PrevInstrExecuted = null;
        }

        //--manage->save const value (string) and varname, if a fctcall or string concatenation is found, push it on the stack and process it
        while(true)
        {
            instrSelectFiles.CurrParamNum++;

            // no more param to process
            if (instrSelectFiles.CurrParamNum >= instrSelectFiles.ListInstrParams.Count)
                // next stage: select files, based on selected and unselected files
                break;

            // get the current param and selector
            InstrBase param = instrSelectFiles.ListInstrParams[instrSelectFiles.CurrParamNum];
            InstrSelectFilesSelector selector = instrSelectFiles.ListFilesSelectors[instrSelectFiles.CurrParamNum];

            // is it a const value (string) or varname?
            if(param.InstrType== InstrType.ConstValue || param.InstrType== InstrType.ObjectName)
            {
                instrSelectFiles.RunTmpListFinalInstrParams.Add(param);
                continue;
            }

            // fct call, string concatenation, the param have to be processed
            ctx.StackInstr.Push(param);
            return true;
        }

        //--list of params have been all decoded, contains only constValue or varname
        for(int i=0; i <instrSelectFiles.RunTmpListFinalInstrParams.Count; i++)
        {
            // get the current param and selector
            InstrBase param = instrSelectFiles.ListInstrParams[i];
            InstrSelectFilesSelector selector = instrSelectFiles.ListFilesSelectors[i];

            // the param is a const value or a varname, get final lsit of filename to process: apply select and unselect filters
            if (!DecodeParam(execResult, progRunVarMgr, instrSelectFiles, param, selector))
                return false;
        }

        // save the instr and remove it from the stack
        ctx.PrevInstrExecuted = instrSelectFiles;
        ctx.StackInstr.Pop();
        return true;
    }

    bool DecodeParam(ExecResult execResult, ProgRunVarMgr progRunVarMgr, InstrSelectFiles instrSelectFiles, InstrBase param, InstrSelectFilesSelector selector)
    {
        InstrConstValue instrConstValue;

        //--1/param is constValue type string? exp: SelectFiles("file.xlsx")
        instrConstValue = param as InstrConstValue;
        if (instrConstValue != null)
        {
            if(!SelectFilesFromStringFilename(execResult, instrSelectFiles, instrConstValue, out List<string> listFilename))
                return false;
        }

        //--2/param is a ObjectName ? exp: SelectFiles(fileName)
        InstrObjectName instrObjectName = param as InstrObjectName;
        if (instrObjectName != null)
        {
            // get the var name, should be defined before, can be a var of var, exp: a=b so return the last var (b in the sample)
            ProgRunVar execVar = progRunVarMgr.FindLastInnerVarByName(instrObjectName.ObjectName);
            if (execVar == null)
            {
                execResult.AddError(ErrorCode.RunInstrVarNotFound, instrObjectName.FirstScriptToken());
                return false;
            }

            //-the value of the var is a string constValue?
            instrConstValue = execVar.Value as InstrConstValue;
            if (instrConstValue != null)
            {
                if (!SelectFilesFromStringFilename(execResult, instrSelectFiles, instrConstValue, out List<string> listFilename))
                    return false;
                return true;
            }
        }

        execResult.AddError(ErrorCode.RunInstrNotManaged, param.FirstScriptToken());
        return true;
    }

    bool SelectFiles(ExecResult execResult, InstrBase instrBase, string filename, out List<string> listFilenameOut)
    {
        listFilenameOut = new List<string>();
        try
        {
            if (string.IsNullOrEmpty(filename))
            {
                execResult.AddError(ErrorCode.RunInstrFilenameWrong, instrBase.FirstScriptToken());
                return false;
            }

            string filepath = Path.GetDirectoryName(filename);
            string fullpath = Path.GetFullPath(filename);
            string files = Path.GetFileName(filename);

            if (string.IsNullOrEmpty(filepath))
                filepath = ".";

            if (!Path.Exists(filepath))
            {
                execResult.AddError(ErrorCode.RunInstrFilePathWrong, instrBase.FirstScriptToken());
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
            execResult.AddError(ErrorCode.RunInstrAccessFileWrong, instrBase.FirstScriptToken(),ex);
            return false;
        }
    }


    bool SelectFilesFromStringFilename(ExecResult execResult, InstrSelectFiles instrSelectFiles,  InstrConstValue instrConstValue, out List<string> listFilename)
    {
        // should be a string
        ValueString valueString = instrConstValue.ValueBase as ValueString;
        if (valueString == null)
        {
            listFilename = new List<string>();
            execResult.AddError(ErrorCode.RunInstrTypeStringExpected, instrConstValue.FirstScriptToken());
            return false;
        }

        if (!SelectFiles(execResult, instrConstValue, StringUtils.RemoveStartEndDoubleQuote(valueString.Val), out listFilename))
            return false;

        // save list of files
        foreach (string filename in listFilename)
            instrSelectFiles.AddFinalFilename(instrConstValue, filename);
        return true;
    }
}
