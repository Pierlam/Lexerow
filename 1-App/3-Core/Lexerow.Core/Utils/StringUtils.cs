using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Utils;

/// <summary>
/// String utilities.
/// </summary>
public class StringUtils
{
    /// <summary>
    /// file -> not a string.
    /// "file.xlsx"  -> it's a string.
    /// 
    /// "file.xlsx  -> error, wrong.
    /// </summary>
    /// <param name="str"></param>
    /// <param name="isString"></param>
    /// <returns></returns>
    public static bool HasStartEndQuotes(string str, out bool isString)
    {
        isString = false;
        if (string.IsNullOrWhiteSpace(str)) return false;
        if (str.StartsWith("\"", StringComparison.OrdinalIgnoreCase) && str.EndsWith("\"", StringComparison.OrdinalIgnoreCase))
        {
            isString = true;
            return true;
        }

        // not a string
        isString = false;
        return true;
    }

    public static string RemoveStartEndDoubleQuote(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return s;
        string str=s.Trim();

        if (str.StartsWith("\"", StringComparison.OrdinalIgnoreCase))
            str=str.Substring(1);

        if(str.EndsWith("\"", StringComparison.OrdinalIgnoreCase))
            str=str.Substring(0, str.Length - 1);

        return str;
    }

}
