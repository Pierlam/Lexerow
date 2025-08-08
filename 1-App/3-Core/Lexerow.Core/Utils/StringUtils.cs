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

}
