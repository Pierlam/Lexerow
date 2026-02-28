using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.Object;
using OpenExcelSdk;

namespace Lexerow.Core.System;

/// <summary>
/// Program execution context.
/// Used during the execution of script program.
/// </summary>
public class ProgExecContext
{
    public Stack<InstrBase> StackInstr { get; private set; } = new Stack<InstrBase>();

    /// <summary>
    /// The previous instruction executed, used for some instructions which need to get the info of the previous instruction, e.g. ForEachRow instruction need to get the excel sheet and row index info from the context which is set by the previous OnExcel instruction.
    /// </summary>
    public InstrBase PrevInstrExecuted { get; set; } = null;

    /// <summary>
    /// The current excel file object to process.
    /// </summary>
    public InstrObjectExcelFile ExcelFileObject { get; set; } = null;

    /// <summary>
    /// The current excel sheet to process.
    /// </summary>
    public ExcelSheet? ExcelSheet { get; set; } = null;

    /// <summary>
    /// The current excel sheet row index to process.
    /// Starting from 1, init to -1 which is not started.
    /// </summary>
    public int RowIndex { get; set; } = -1;

    /// <summary>
    /// list of selected excel filename to process
    /// </summary>
    public List<ObjectSelectedFile> ListObjectSelectedFiles { get; private set; } = new List<ObjectSelectedFile>();
}