using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.ProgRun;
public class ProgramRunnerContext
{
    public Stack<InstrBase> StackInstr { get; private set; } = new Stack<InstrBase>();

    public InstrBase PrevInstrExecuted { get; set;} = null;

    /// <summary>
    /// Init to -1 which is not started.
    /// Used in execution.
    /// base0.
    /// </summary>
    public int FileToProcessNum { get; set; } = -1;

    /// <summary>
    /// The current excel file object to process.
    /// </summary>
    public InstrExcelFileObject ExcelFileObject { get; set; } = null;

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
    public List<InstrSelectedFilename> ListSelectedFilename { get; private set; } = new List<InstrSelectedFilename>();

}
