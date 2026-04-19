using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Diag;


///// <summary>
///// 2 two ways to use the message item:
///// 1/ Get message by operation, stage and result. 
///// 2/ Get message by operation and error code.
///// </summary>
//public class MessageItem 
//{
//    public string Operation { get; set; }
//    public ActivityLogStage Stage { get; set; }

//    public ActivityLogResult Result { get; set; }= ActivityLogResult.Ok;

//    public ErrorCode ErrorCode { get; set; } = ErrorCode.Ok;
//    public string Message { get; set; }

//    public bool UseParam { get; set; } 

//    /// <summary>
//    /// Case result is ok.
//    /// </summary>
//    /// <param name="operation"></param>
//    /// <param name="stage"></param>
//    /// <param name="msg"></param>
//    /// <param name="useParam"></param>
//    public MessageItem(string operation, ActivityLogStage stage, string msg, bool useParam)
//    {
//        Operation = operation;
//        Stage = stage;
//        Message = msg;
//        UseParam = useParam;
//    }

//    /// <summary>
//    /// Case result is ok.
//    /// </summary>
//    /// <param name="operation"></param>
//    /// <param name="stage"></param>
//    /// <param name="msg"></param>
//    /// <param name="useParam"></param>
//    public MessageItem(string operation, ActivityLogStage stage, ActivityLogResult result, string msg, bool useParam)
//    {
//        Operation = operation;
//        Stage = stage;
//        Result=result;
//        Message = msg;
//        UseParam = useParam;
//    }

//    public MessageItem(string operation, ErrorCode errorCode, string msg, bool useParam)
//    {
//        Operation = operation;
//        Stage = ActivityLogStage.Start;
//        ErrorCode = errorCode;
//        Message = msg;
//        UseParam = useParam;
//    }

//}

public class MessageBuilder
{
    //List<MessageItem> _listMsg = new List<MessageItem>();
    UserMessageRegistry _listMsg = new UserMessageRegistry();

    public MessageBuilder()
    {
        // build the registry of messages
        BuildListMsg();
    }

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

        ResultError result = log.Error;
        if(result==null)result= log.Warning;

        string msgFromRegistry = GetMsg(log.Operation, log.Result, result, log.Param, log.Param2, log.Param3, log.Param4, log.Param5);
        if(!string.IsNullOrEmpty(msgFromRegistry))
        {
            return msg += msgFromRegistry;
        }

        return string.Empty ;
    }

    /// <summary>
    /// Build the list of user/human readable messages.
    /// </summary>
    void BuildListMsg()
    {
        _listMsg.Add("LexerowCore.LoadExecScript.Start", "Start process script {0}.");
        _listMsg.Add("LexerowCore.LoadExecScript.End","End Process script {0}, elapsed time: {1}, success.");
        _listMsg.AddErrorCode("LexerowCore.LoadExecScript", ErrorCode.FileNameNullOrEmpty, "End Process script failed, script file name is null or empty.");
        _listMsg.AddError("LexerowCore.LoadExecScript", "End Process script {0} failed.");


        _listMsg.Add("LexerowCore.LoadScriptFromFile.Start", "Start load script from file {0}.");
        _listMsg.Add("LexerowCore.LoadScriptFromFile.End", "End load script, {0} lines found, success.");
        _listMsg.AddErrorCode("LexerowCore.LoadScriptFromFile", ErrorCode.FileNotFound, "End load script failed, file {0} not found.");
        _listMsg.AddError("LexerowCore.LoadScriptFromFile", "End load script from file {0} failed.");

        // ScriptCompiler.CompileScript
        _listMsg.Add("ScriptCompiler.CompileScript.Start", "Start compile script");
        _listMsg.Add("ScriptCompiler.CompileScript.End", "End compile script, {0} instruction(s) generated, success.");

        _listMsg.AddErrorCode("ScriptCompiler.CompileScript", ErrorCode.ParserTokenNotExpected, "End compile script failed, L{0}:C{1}, token '{2}' not expected.");        
        _listMsg.AddErrorCode("ScriptCompiler.CompileScript", ErrorCode.ParserValueStringExpected, "End compile script failed, L{0}:C{1}, token '{2}' should be a string value.");
        _listMsg.AddErrorCode("ScriptCompiler.CompileScript", ErrorCode.ParserExpressionWrong, "End compile script failed, L{0}:C{1}, wrong expression on token '{2}'.");
        _listMsg.AddErrorCode("ScriptCompiler.CompileScript", ErrorCode.ParserCompExprWrong, "End compile script failed, L{0}:C{1}, wrong comparison expression on token '{2}'.");
        _listMsg.AddErrorCode("ScriptCompiler.CompileScript", ErrorCode.ParserSepComparatorWrong, "End compile script failed, L{0}:C{1}, token '{2}' is a wrong separator comparison.");
        _listMsg.AddErrorCode("ScriptCompiler.CompileScript", ErrorCode.ParserBoolExprWrong, "End compile script failed, L{0}:C{1}, wrong boolean expression on token '{2}'.");

        

        // ProgramExecutor.Exec
        _listMsg.Add("ProgramExecutor.Exec.Start", "Start execute script.");
        _listMsg.Add("ProgramExecutor.Exec.End", "End execute script, success.");
        _listMsg.AddError("ProgramExecutor.Exec", "End execute script failed.");


        // InstrExecutor.ExecInstr
        _listMsg.Add("InstrExecutor.ExecInstr.Start", "Start execute instruction {0}.");
        _listMsg.Add("InstrExecutor.ExecInstr.End", "End execute instruction {0}, success.");
        _listMsg.Add("InstrOnExcelExecutor.ExecInstrOnExcel.ProcessFile", "Process Excel File {0}.");

        _listMsg.AddErrorCode("InstrExecutor.ExecInstr", ErrorCode.ExecNoFileSelected, "End execute script failed, L{0}:C{1}, no file selected match {2}, nothing to process.");

    }

    string GetMsg(string operation, ActivityLogResult result, ResultError error, string param, string param2, string param3, string param4, string param5)
    {
        return _listMsg.GetMsg(operation, result, error, param, param2, param3, param4, param5);
    }
}
