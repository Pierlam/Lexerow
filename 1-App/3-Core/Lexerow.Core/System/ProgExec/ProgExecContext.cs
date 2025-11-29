using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.Func;
using Lexerow.Core.System.InstrDef.Object;

namespace Lexerow.Core.System;

/// <summary>
/// Program execution context.
/// Used during the execution of script program.
/// </summary>
public class ProgExecContext
{
    public Stack<InstrBase> StackInstr { get; private set; } = new Stack<InstrBase>();

    public InstrBase PrevInstrExecuted { get; set; } = null;

    /// <summary>
    /// Init to -1 which is not started.
    /// Used in execution.
    /// base0.
    /// </summary>
    public int FileToProcessNum { get; set; } = -1;

    /// <summary>
    /// The current excel file object to process.
    /// </summary>
    public InstrObjectExcelFile ExcelFileObject { get; set; } = null;

    /// <summary>
    /// The current excel sheet to process.
    /// </summary>
    public IExcelSheet? ExcelSheet { get; set; } = null;

    /// <summary>
    /// The current excel sheet row num to process.
    /// </summary>
    public int RowNum { get; set; } = -1;

    /// <summary>
    /// list of selected excel filename to process
    /// </summary>
    public List<InstrFuncSelectedFilename> ListSelectedFilename { get; private set; } = new List<InstrFuncSelectedFilename>();
}