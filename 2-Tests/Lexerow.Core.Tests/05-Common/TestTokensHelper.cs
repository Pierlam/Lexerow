using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.Tests._05_Common;

internal class TestTokensHelper
{
    public static bool TestString(ScriptLineTokens lineTokens, int idx, string value)
    {
        return Test(lineTokens, idx, value, ScriptTokenType.String);
    }

    /// <summary>
    ///  a system variable . Exp: $DateFormat
    /// </summary>
    /// <param name="lineTokens"></param>
    /// <param name="idx"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public static bool TestSystName(ScriptLineTokens lineTokens, int idx, string value)
    {
        return Test(lineTokens, idx, value, ScriptTokenType.SystName);
    }

    /// <summary>
    ///  a variable name , or fct, class,...
    /// </summary>
    /// <param name="lineTokens"></param>
    /// <param name="idx"></param>
    /// <param name="value"></param>
    /// <returns></returns>
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
        if (lineTokens.ListScriptToken[idx].ValueInt == value) return true;
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