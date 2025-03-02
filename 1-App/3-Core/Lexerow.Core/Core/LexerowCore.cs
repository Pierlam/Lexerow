using Lexerow.Core.ExcelLayer;
using Lexerow.Core.Logger;
using Lexerow.Core.System.Excel;
using Microsoft.Extensions.Logging;

namespace Lexerow.Core;
public class LexerowCore
{
    ILoggerFactory _loggerFactory;

    IExcelProcessor _excelProcessor;

    CoreData _coreData=new CoreData();

    /// <summary>
    /// Constructor.
    /// </summary>
    public LexerowCore()
    {
        _loggerFactory = new InternalLoggerFactory();
        _excelProcessor= new ExcelProcessorNpoi();

        Builder= new CoreBuilder(_loggerFactory, _coreData);
        Exec = new Exec(_loggerFactory, _coreData, _excelProcessor);
    }

    public LexerowCore(ILoggerFactory loggerFactory)
    {
        _loggerFactory= loggerFactory;
    }

    /// <summary>
    /// Build instruction, expression,...
    /// </summary>
    public CoreBuilder Builder { get; private set; }

    /// <summary>
    /// Execute instructions.
    /// </summary>
    public Exec Exec { get; private set; }
}
