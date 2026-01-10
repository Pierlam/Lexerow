using Lexerow.Core.System.ActivLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Diag;

/// <summary>
/// to do some internal diagnostics
/// </summary>
public class Diagnostics
{
    private IActivityLogger _logger;
    StreamWriter _swTxt = null;

    StreamWriter _swCsv = null;

    string _csvSep = ";";

    public Diagnostics(IActivityLogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// is save log to a file is actice?
    /// </summary>
    public bool IsLogToConsoleActive { get; set; } = false;

    /// <summary>
    /// is save log to a text file is actice?
    /// </summary>
    public bool IsSaveLogTxtActive { get; set; } = false;

    /// <summary>
    /// is save log to a csv file is actice?
    /// </summary>
    public bool IsSaveLogCsvActive { get; set; } = false;

    public void LogToConsole(bool allow)
    {
        IsLogToConsoleActive = true;
    }

    /// <summary>
    /// Stop the log activity.
    /// </summary>
    public void SetLogLevelOff()
    {
        _logger.ActiveLevel = ActivityLogLevel.Off;
    }

    /// <summary>
    /// Autorize only the highest log lelve: Info
    /// </summary>
    public void SetLogLevelInfo()
    {
        _logger.ActiveLevel = ActivityLogLevel.Info;
    }

    /// <summary>
    /// Authorize the activity level Info and Debug.
    /// </summary>
    public void SetLogLevelDebug()
    {
        _logger.ActiveLevel = ActivityLogLevel.Debug;
    }

    /// <summary>
    /// Authorize all activity level Info, Debug and Trace.
    /// </summary>
    public void SetLogLevelTrace()
    {
        _logger.ActiveLevel = ActivityLogLevel.Trace;
    }

    /// <summary>
    /// Save te log into text and csv files if its activated.
    /// </summary>
    /// <param name="e"></param>
    public void SaveLog(ActivityLog e)
    {
        if (_swTxt != null) _swTxt.WriteLine(BuildLogLine(e));
        if (_swCsv != null) _swCsv.WriteLine(BuildLogLineSep(e, _csvSep));
    }

    public void LogToConsole(ActivityLog e)
    {
        if (!IsLogToConsoleActive) return;
        Console.WriteLine(BuildLogLine(e));
    }

    /// <summary>
    /// Save logs into a text file.
    /// </summary>
    /// <param name="filename"></param>
    /// <returns></returns>
    public bool SaveLogTxt(string filename)
    {
        _swTxt = CreateTextFile(filename);
        if(_swTxt==null) return false;
        IsSaveLogTxtActive = true;
        return true;
    }

    public bool SaveLogCsv(string filename)
    {
        _swCsv = CreateTextFile(filename);
        if (_swCsv == null) return false;
        IsSaveLogCsvActive = true;
        return true;
    }

    public void CloseLogs()
    {
        if(_swTxt!=null) _swTxt.Close();
        if (_swCsv != null) _swCsv.Close();

        _swTxt = null;
        _swCsv = null;
    }

    string BuildLogLine(ActivityLog e)
    {
        string s = string.Empty;

        s = e.When.ToString("dd/MM/yyyy HH:mm:ss fff");

        s += " " + e.Level + " " + e.Module + " " + e.Stage + " " + e.Result + " " + e.Operation + " " + e.Msg;
        return s;
    }

    string BuildLogLineSep(ActivityLog e, string sep)
    {
        string s = string.Empty;

        s = e.When.ToString("dd/MM/yyyy HH:mm:ss fff");

        s += sep + e.Level + sep + e.Module + sep + e.Stage + sep + e.Result + sep + e.Operation + sep + e.Msg;
        return s;
    }

    StreamWriter CreateTextFile(string filename)
    {
        try
        {
            // remove the preivous file
            if (File.Exists(filename))
            {
                File.Delete(filename);
            }

            // Create a file to write to
            return  File.CreateText(filename);
        }
        catch (Exception ex)
        {
            return null;
        }
    }

}
