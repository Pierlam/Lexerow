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

        ProgBuilder= new ProgBuilder(_coreData);
        Exec = new Exec(_coreData, _excelProcessor);

        _scriptLoader= new ScriptLoader();
        _scriptCompilator = new ScriptCompilator(_coreData);

        // create the default program, becomes the current one
        ProgramInstr programInstr= new ProgramInstr("Default");
        programInstr.IsDefault = true;
        _coreData.ListProgram.Add(programInstr);
        _coreData.CurrProgramInstr= programInstr;
    }

    /// <summary>
    /// Build program by adding instructions.
    /// Don't manage script/source code.
    /// </summary>
    public ProgBuilder ProgBuilder { get; private set; }

    /// <summary>
    /// Execute instructions.
    /// </summary>
    public Exec Exec { get; private set; }

    public ExecResult CompileProgram()
    {
        return Exec.CompileProgram();
    }

    public ExecResult CompileProgram(string programName)
    {
        return Exec.CompileProgram(programName);
    }

    public ExecResult ExecuteProgram()
    {
        return Exec.ExecuteProgram();
    }

    public ExecResult ExecuteProgram(string programName)
    {
        return Exec.ExecuteProgram(programName);
    }

    /// <summary>
    /// Load a script from a text file and compile it.
    /// </summary>
    /// <param name="progName"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public ExecResult LoadScriptFromFile(string programName, string fileName)
    {
        ExecResult execResult = new ExecResult();

        // check that the name is not already used by another program
        if (string.IsNullOrWhiteSpace(programName))
        {
            if (programName == null) programName = string.Empty;
            execResult.AddError(new ExecResultError(ErrorCode.ProgramWrongName, programName));
            return execResult;
        }

        // any program should have the same name
        ProgramInstr program = _coreData.GetProgramByName(programName);
        if (program != null)
        {
            execResult.AddError(new ExecResultError(ErrorCode.ProgramNameAlreadyUsed, programName));
            return execResult;
        }

        _scriptLoader.LoadScriptFromFile(execResult, fileName, out SourceScript sourceScript);
        if(!execResult.Result)
            return execResult;

        // compile the script,  generate instructions
        _scriptCompilator.CompileScript(execResult, sourceScript, out List<InstrBase> listInstr);
        if (!execResult.Result)
            return execResult;

        // create the program
        ProgramInstr programInstr = new ProgramInstr(programName, fileName, sourceScript, listInstr);

        // save it
        _coreData.ListProgram.Add(programInstr);

        return execResult;
    }

    //ExecResult Core.LoadScriptFromLines(string name, List<string> scriptLines)

    //ExecResult Core.LoadExecScriptFromFile(string name, string filename)

    //ExecResult Core.LoadExecScriptFromLines(string name, List<string> scriptLines)


    public Action<AppTrace> AppTraceEvent 
    {
        get { return AppTraceEvent; } 
        set { 
            ProgBuilder.AppTraceEvent = value;
            Exec.AppTraceEvent = value;
        }
    }
}
