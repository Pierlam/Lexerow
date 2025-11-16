namespace Lexerow.Core.Tests.Common;

public abstract class BaseTests
{
    /// <summary>
    /// Path where excel files are saved.
    /// </summary>
    public string PathExcelFilesExec = @".\10-ExcelFiles\";

    /// <summary>
    /// Path where script are saved.
    /// </summary>
    public string PathScriptFiles = @"..\..\..\15-Scripts\";

    /// <summary>
    /// add double quote at the start of the string and also and the end
    /// exp: file.xlsx  -> "file.xlsx"
    /// </summary>
    /// <param name="s"></param>
    /// <returns></returns>
    public string AddDblQuote(string s)
    {
        return "\"" + s + "\"";
    }
}