using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using NPOI.OpenXmlFormats.Spreadsheet;
using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts.SyntaxAnalyze;
internal class SetVarDecoder
{
    /// <summary>
    /// Manage case  equal char is found, as a set variable instruction.
    /// exp:
    ///   file=                 implemented
	///   A.Cell= 	            TODO
	///   A.Cell.BgColor =      TODO
	///   excelOut.B.Cell=      TODO
	///   Sheet[1].B.Cell=      TODO?
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stkItems"></param>
    /// <param name="scriptToken"></param>
    /// <param name="listExecTok"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool ProcessSetVarEqualChar(ExecResult execResult, List<InstrObjectName> listVar, Stack<InstrBase> stkItems, ScriptToken scriptToken, List<InstrBase> listExecTok, out bool isToken)
    {
        isToken = false;

        // is the script token the equal char?
        if(!scriptToken.Value.Equals("=",StringComparison.InvariantCultureIgnoreCase))
            // not the equal char, bye without error
            return true;

        // the stack contains nothing, strange  =blabla
        if(stkItems.Count == 0)
        {
            execResult.AddError(new ExecResultError(ErrorCode.SyntaxAnalyzerTokenNotExpected, "="));
            return false;
        }

        isToken = true;

        // the stack contains many tokens, bye without error
        if (stkItems.Count > 1)
            // TODO: manage this case! can be cases A.Cell, A.Cell.BgColor, ...?
            return true;

        // the stack contains one item which is an Object name, exp: file=
        if(stkItems.Peek().InstrType != InstrType.ObjectName)
        {
            execResult.AddError(new ExecResultError(ErrorCode.SyntaxAnalyzerTokenNotExpected, stkItems.Peek().FirstScriptToken().Value));
            return false;
        }

        InstrObjectName instrObjectName = stkItems.Peek() as InstrObjectName;
        InstrSetVar instrSetVar = new InstrSetVar(instrObjectName.FirstScriptToken());
        instrSetVar.InstrLeft= instrObjectName;

        instrSetVar.ListScriptToken.Add(scriptToken);

        // remove the objectName from the stack and push the SetVar instr
        stkItems.Pop();
        stkItems.Push(instrSetVar);

        // save the var
        listVar.Add(instrObjectName);

        return true;
    }
}
