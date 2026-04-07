using Lexerow.Core.System.ActivLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Diag;

public class MessageBuilder
{
    public string BuildMsg(ActivityLog log)
    {
        string msg = $"{log.When:yyyy-MM-dd HH:mm:ss.fff} ";

        if(log.Result == ActivityLogResult.Error)
        {
            msg += "ERR ";
        }
        else if (log.Result == ActivityLogResult.Warning)
        {
            msg += "WRN ";
        }
        else
        {
            msg += "INF ";
        }

        if (log.Operation.Equals("LexerowCore.LoadExecScript") && log.Stage== ActivityLogStage.Start)
        {
            return msg += $"Start process script {log.Param}.";
        }
        if (log.Operation.Equals("LexerowCore.LoadExecScript") && log.Stage == ActivityLogStage.End)
        {
            // success or error
            if (log.Result == ActivityLogResult.Error)
            {
                int errorCount = 0;
                return msg += $"Process script {log.Param} finished with error(s).";
            }
            else
            {
                return msg += $"End Process script {log.Param} finished with success.";
            }
        }

        if (log.Operation.Equals("LexerowCore.LoadScriptFromFile"))
        {
            if(log.Stage== ActivityLogStage.Start) 
                return msg+= $"Start load script from file {log.Param}.";
            if (log.Stage == ActivityLogStage.End)
            {
                if (log.Result == ActivityLogResult.Error)
                    return msg += $"End load script from file {log.Param} finished with error.";
                else
                    return msg += $"End load script from file {log.Param} finished with success.";
            }
        }

        if(log.Operation.Equals("LexerowCore.ExecuteScript"))
        {
            if (log.Stage == ActivityLogStage.Start)
                return msg += $"Start Execute script.";
            if (log.Stage == ActivityLogStage.End)
            {
                if (log.Result == ActivityLogResult.Error)
                    return msg += $"End execute script finished with error(s).";
                else
                    return msg += $"End execute script finished with success.";
            }
        }
        return string.Empty ;
    }
}
