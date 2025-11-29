using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.InstrFuncDef;
using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.InstrProgExec;
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
        // check the year instr, can be a Value, a var or a fctcall
        //if(instrFuncDate.InstrDay)

        //InstrObjectDate

        // manage cases when instr Left and/or Right have to be executed: fct call
        //if (instrFuncDate.LastInstrExecuted > 0)
        return false;
    }
}
