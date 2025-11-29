using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgExec;
public class InstrFuncDateExecutor
{
    IActivityLogger _activityLogger;

    public InstrFuncDateExecutor(IActivityLogger activityLogger)
    {
        _activityLogger = activityLogger;
    }

    /// <summary>
    /// Execute the function Date.
    /// exp: Date(year, month, day)
    /// Generate a date object.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="ctx"></param>
    /// <param name="progExecVarMgr"></param>
    /// <param name="instrFuncDate"></param>
    /// <returns></returns>
    public bool ExecFuncDate(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrFuncDate instrFuncDate)
    {
        //InstrObjectDate
        //InstrExcelFileObject
        //InstrObjectName
    }
}
