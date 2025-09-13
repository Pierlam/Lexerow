using Lexerow.Core.ExcelLayer;
using Lexerow.Core.Scripts;
using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using Lexerow.Core.System.Excel;
using NPOI.HPSF;
using NPOI.SS.Formula.Functions;

namespace Lexerow.Core;

/// <summary>
/// Lexerow core backend application.
/// </summary>
public class LexerowCore
{
    IExcelProcessor _excelProcessor;

    CoreData _coreData=new CoreData();

    ScriptLoader _scriptLoader;

    ScriptCompilator _scriptCompilator;


    /// <summary>
    /// Constructor.
    /// </summary>
    public LexerowCore()
    {
        _excelProcessor= new ExcelProcessorNpoi();

        Exec = new Exec(_coreData, _excelProcessor);

        _scriptLoader= new ScriptLoader();
        _scriptCompilator = new ScriptCompilator(_coreData);
    }

    /// <summary>
    /// Build program by adding instructions.
    /// Don't manage script/source code.
    /// TODO: REMOVE IT?
    /// </summary>
    //public ProgBuilder ProgBuilder { get; private set; }

    /// <summary>
    /// Execute instructions.
    /// </summary>
    public Exec Exec { get; private set; }

    //public ExecResult CompileProgram()
    //{
    //    return Exec.CompileProgram();
    //}

    /// <summary>
    /// TODO: REMOVE IT?
    /// </summary>
    /// <returns></returns>
    public ExecResult ExecuteProgram()
    {
        return Exec.ExecuteProgram();
    }

    //public ExecResult ExecuteProgram(string programName)
    //{
    //    return Exec.ExecuteProgram(programName);
    //}

    /// <summary>
    /// Load a script from a text file and compile it.
    /// </summary>
    /// <param name="progName"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public ExecResult LoadScriptFromFile(string scriptName, string fileName)
    {
        ExecResult execResult = new ExecResult();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(scriptName))
        {
            if (scriptName == null) scriptName = string.Empty;
            execResult.AddError(new ExecResultError(ErrorCode.ProgramWrongName, scriptName));
            return execResult;
        }

        // any program should have the same name
        ProgramScript program = _coreData.GetProgramByName(scriptName);
        if (program != null)
        {
            execResult.AddError(new ExecResultError(ErrorCode.ProgramNameAlreadyUsed, scriptName));
            return execResult;
        }

        _scriptLoader.LoadScriptFromFile(execResult, scriptName, fileName, out Script script);
        if(!execResult.Result)
            return execResult;

        // compile the script,  generate instructions
        _scriptCompilator.CompileScript(execResult, script, out List<ExecTokBase> listInstr);
        if (!execResult.Result)
            return execResult;

        // create the program
        ProgramScript programInstr = new ProgramScript(script, listInstr);

        // save it
        _coreData.ListProgram.Add(programInstr);

        return execResult;
    }

    public ExecResult LoadScriptFromLines(string scriptName, List<string> scriptLines)
    {
        // TODO:
        return null;
    }

    //ExecResult LoadExecScriptFromFile(string name, string filename)

    //ExecResult LoadExecScriptFromLines(string name, List<string> scriptLines)

    public ExecResult ExecuteScript(string name)
    {
        throw new Exception("todo");
    }

    Action<AppTrace> _appTraceEvent;

    public Action<AppTrace> AppTraceEvent 
    {
        get { return _appTraceEvent; } 
        set { 
            //ProgBuilder.AppTraceEvent = value;
            Exec.AppTraceEvent = value;
        }
    }
}
