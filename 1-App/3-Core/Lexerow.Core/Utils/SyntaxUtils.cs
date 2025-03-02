using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Utils;
public class SyntaxUtils
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
}
