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
    /// The current excel file object.
    /// </summary>
    public InstrExcelFileObject ExcelFileObject { get; set; } = null;

    /// <summary>
    /// Used in execution.
    /// base0.
    /// </summary>
    //public int SheetToProcessNum { get; set; } = -1;

    public IExcelSheet? ExcelSheet { get; set; } = null;

    public int RowNum { get; set; } = -1;

    /// <summary>
    /// Final list of excel file.
    /// Used in execution.
    /// Excel file object can be loaded at the last time.
    /// </summary>
    public List<InstrExcelFileObject> ListExcelFileObject { get; set; } = new List<InstrExcelFileObject>();

}
