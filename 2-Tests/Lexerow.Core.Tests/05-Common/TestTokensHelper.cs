using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests._05_Common;
internal class TestTokensHelper
{
    public static bool TestName(ScriptLineTokens lineTokens, int idx, string value)
    {
        return Test(lineTokens, idx, value, ScriptTokenType.Name);
    }

    public static bool TestSep(ScriptLineTokens lineTokens, int idx, string value)
    {
        return Test(lineTokens, idx, value, ScriptTokenType.Separator);
    }

    public static bool TestNumberInt(ScriptLineTokens lineTokens, int idx, int value)
    {
        if (lineTokens.ListScriptToken[idx] == null) return false;
        if (lineTokens.ListScriptToken[idx].ValueInt== value) return true;
        if (lineTokens.ListScriptToken[5].ScriptTokenType == ScriptTokenType.Integer) return true;

        return false;
    }

    public static bool TestNumberDouble(ScriptLineTokens lineTokens, int idx, double value)
    {
        if (lineTokens.ListScriptToken[idx] == null) return false;
        if (lineTokens.ListScriptToken[idx].ValueInt == value) return true;
        if (lineTokens.ListScriptToken[5].ScriptTokenType == ScriptTokenType.Double) return true;

        return false;
    }

    public static bool Test(ScriptLineTokens lineTokens, int idx, string value, ScriptTokenType type)
    {
        if (lineTokens.ListScriptToken[idx] == null) return false;
        if (lineTokens.ListScriptToken[idx].Value == value) return true;
        if (lineTokens.ListScriptToken[5].ScriptTokenType == type) return true;

        return false;
    }
}
