using Lexerow.Core.System.ActivLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// 2 two ways to use the message item:
/// 1/ Get message by operation, stage and result. 
/// 2/ Get message by operation and error code.
/// </summary>
public class MessageItem
{
    public string Operation { get; set; }
    public ActivityLogStage Stage { get; set; }

    public ActivityLogResult Result { get; set; } = ActivityLogResult.Ok;

    public ErrorCode ErrorCode { get; set; } = ErrorCode.Ok;
    public string Message { get; set; }

    public bool UseParam { get; set; }

    /// <summary>
    /// Case result is ok.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="stage"></param>
    /// <param name="msg"></param>
    /// <param name="useParam"></param>
    public MessageItem(string operation, ActivityLogStage stage, string msg, bool useParam)
    {
        Operation = operation;
        Stage = stage;
        Message = msg;
        UseParam = useParam;
    }

    /// <summary>
    /// Case result is ok.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="stage"></param>
    /// <param name="msg"></param>
    /// <param name="useParam"></param>
    public MessageItem(string operation, ActivityLogStage stage, ActivityLogResult result, string msg, bool useParam)
    {
        Operation = operation;
        Stage = stage;
        Result = result;
        Message = msg;
        UseParam = useParam;
    }

    public MessageItem(string operation, ErrorCode errorCode, string msg, bool useParam)
    {
        Operation = operation;
        Stage = ActivityLogStage.Start;
        ErrorCode = errorCode;
        Message = msg;
        UseParam = useParam;
    }

}

public class UserMessageRegistry
{
    List<MessageItem> _listMsg = new List<MessageItem>();

    public void AddStart(string operation, string msg)
    {
        Add(operation, ActivityLogStage.Start, msg);
    }

    public void AddOnGoing(string operation, string msg)
    {
        Add(operation, ActivityLogStage.OnGoing, msg);
    }

    public void AddEnd(string operation, string msg)
    {
        Add(operation, ActivityLogStage.End, msg);
    }

    public void AddEndError(string operation, string msg)
    {
        Add(operation, ActivityLogStage.End, ActivityLogResult.Error, msg);
    }

    public void Add(string operation, ActivityLogStage stage, string msg)
    {
        bool useParam = false;

        if (string.IsNullOrEmpty(operation)) operation = "(null)";

        if (msg.Contains("{0}"))useParam = true;
                    
        _listMsg.Add(new MessageItem(operation, stage, msg, useParam));
    }

    public void Add(string operation, ErrorCode errorCode, string msg)
    {
        bool useParam = false;

        if (string.IsNullOrEmpty(operation)) operation = "(null)";

        if (msg.Contains("{0}")) useParam = true;
        _listMsg.Add(new MessageItem(operation, errorCode, msg, useParam));
    }

    public void Add(string operation, ActivityLogStage stage, ActivityLogResult result, string msg)
    {
        bool useParam = false;

        if (string.IsNullOrEmpty(operation)) operation = "(null)";

        if (msg.Contains("{0}")) useParam = true;
        _listMsg.Add(new MessageItem(operation, stage, result, msg, useParam));
    }

    /// <summary>
    /// Gte the message based on operation, stage, result and error code. 
    /// If the message with error code exist, it will have higher priority than the message with stage and result.
    /// </summary>
    /// <param name="operation"></param>
    /// <param name="stage"></param>
    /// <param name="result"></param>
    /// <param name="error"></param>
    /// <param name="param"></param>
    /// <param name="param2"></param>
    /// <returns></returns>
    public string GetMsg(string operation, ActivityLogStage stage, ActivityLogResult result, ResultError error, string param, string param2)
    {
        MessageItem item = null;

        // get message based on error code if exist, otherwise get message based on stage and result
        if(error != null)
            item = _listMsg.FirstOrDefault(x => x.Operation == operation && x.ErrorCode == error.ErrorCode);

        if (item == null)
            item = _listMsg.FirstOrDefault(x => x.Operation == operation && x.Stage == stage && x.Result == result);

        if (item == null)
            return string.Empty;

        if (item.UseParam)
        {
            if (string.IsNullOrEmpty(param))
                param = string.Empty;
            if (string.IsNullOrEmpty(param2))
                param2 = string.Empty;

            // use param in log to build the message
            item.Message = string.Format(item.Message, param, param2);
        }
        return item.Message;
    }

}
