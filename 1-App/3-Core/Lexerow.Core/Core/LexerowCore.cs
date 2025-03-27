using Lexerow.Core.ExcelLayer;
using Lexerow.Core.Logger;
using Lexerow.Core.System;
using Lexerow.Core.System.Excel;
using Microsoft.Extensions.Logging;

namespace Lexerow.Core;
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

        Builder= new CoreBuilder(_coreData);
        Exec = new Exec(_coreData, _excelProcessor);
    }

    /// <summary>
    /// Build instruction, expression,...
    /// </summary>
    public CoreBuilder Builder { get; private set; }

    /// <summary>
    /// Execute instructions.
    /// </summary>
    public Exec Exec { get; private set; }

    public Action<AppTrace> AppTraceEvent 
    {
        get { return AppTraceEvent; } 
        set { 
            Builder.AppTraceEvent = value;
            Exec.AppTraceEvent = value;
        }
    }
}
