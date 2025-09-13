using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using NPOI.OpenXmlFormats.Spreadsheet;
using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts;
internal class SetVarDecoder
{
    public static bool ProcessSetVarEqualChar(ExecResult execResult, Stack<SyntaxAnalyserItem> stkItems, ScriptToken scriptToken, List<ExecTokBase> listExecTok, out bool isToken)
    {
        isToken = false;

        // is the script token the equal char?
        if(!scriptToken.Value.Equals("=",StringComparison.InvariantCultureIgnoreCase))
            // not the equla char, bye without error
            return true;

        // the stack contains nothing, strange  =blabla
        if(stkItems.Count == 0)
        {
            execResult.AddError(new ExecResultError(ErrorCode.SyntaxAnalyzerTokenNotExpected, "="));
            return false;
        }

        // the stack contains many token, bye without error
        if (stkItems.Count > 1)
            return true;

        // the stacked token should be a name

        // it's a basic setVar instructions

        // TODO:
        return false;
    }
}
