using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Utils;

public class ItemsCheckUtils
{
    public static bool CheckIdSyntax(string id)
    {
        if (string.IsNullOrWhiteSpace(id)) return false;

        // the first char can't be a digit
        if (char.IsDigit(id[0])) return false;

        foreach (char c in id)
        {
            // not a letter, not a digit and not underscore
            if (!char.IsLetterOrDigit(c) && c != '_') return false;
        }

        return true;
    }

    /// <summary>
    /// Checks items: not null, not empty, and unique.
    /// </summary>
    /// <param name="listItems"></param>
    /// <returns></returns>
    public static bool CheckItemsUnique(List<string> listItems)
    {
        int idx = 0;
        foreach (string item in listItems)
        {           
            if(string.IsNullOrWhiteSpace(item)) return false;
            int idxInner = 0;
            foreach (string i in listItems)
            {
                if (idx!= idxInner && item.Equals(i, StringComparison.InvariantCultureIgnoreCase)) return false;
                idxInner++;
            }

            idx++;
        }
        return true;
    }
}
