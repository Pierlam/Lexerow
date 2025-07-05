using Lexerow.Core.ExcelLayer;
using Lexerow.Core.System;
using Lexerow.Core.System.Excel;

namespace Lexerow.Core;

/// <summary>
/// Lexerow core backend application.
/// </summary>
public class LexerowCore
{
    IExcelProcessor _excelProcessor;

    CoreData _coreData=new CoreData();

    /// <summary>
    /// Constructor.
    /// </summary>
    public LexerowCore()
    {
        _excelProcessor= new ExcelProcessorNpoi();

        ProgBuilder= new ProgBuilder(_coreData);
        Exec = new Exec(_coreData, _excelProcessor);

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

    public Action<AppTrace> AppTraceEvent 
    {
        get { return AppTraceEvent; } 
        set { 
            ProgBuilder.AppTraceEvent = value;
            Exec.AppTraceEvent = value;
        }
    }
}
