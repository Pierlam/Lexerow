using Lexerow.Core.System;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.ScriptCompile.Parse;

internal class SetVarParser
{

    /// <summary>
    /// The current token is equal =, parse the SetVar instruction.
    /// !To execute after ComparisonParser.ParseCompOperator()
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool ProcessSetVar(Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, out bool isToken)
    {
        isToken = false;

        // is the script token the equal char?
        if (!scriptToken.Value.Equals("=", StringComparison.InvariantCultureIgnoreCase))
            // not the equal char, bye without error
            return true;

        // the stack contains nothing, exp: =blabla
        if (stackInstr.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        isToken = true;

        InstrBase instr = stackInstr.Pop();

        // the instr before saved in the stack should be the var name to set
        InstrNameObject instrObjectName = instr as InstrNameObject;
        if (instrObjectName != null) 
        {
            InstrSetVar instrSetVar = new InstrSetVar(instrObjectName.FirstScriptToken());
            instrSetVar.InstrLeft = instrObjectName;
            instrSetVar.ListScriptToken.Add(scriptToken);

            // save the var if it does not exists yet
            if(listVar.FirstOrDefault(v=>v.Name.Equals(instrObjectName.Name, StringComparison.InvariantCultureIgnoreCase))==null)
                listVar.Add(instrObjectName);
            
            stackInstr.Push(instrSetVar);
            return true;
        }

        InstrColCellFunc instrColCellFunc= instr as InstrColCellFunc;
        if(instrColCellFunc != null)
        {
            // now process the SetVar instr
            InstrSetVar instrSetVar = new InstrSetVar(instrColCellFunc.FirstScriptToken());
            instrSetVar.InstrLeft = instrColCellFunc;
            instrSetVar.ListScriptToken.Add(scriptToken);

            stackInstr.Push(instrSetVar);
            return true;
        }

        result.AddError(ErrorCode.ParserTokenNotExpected, instr.FirstScriptToken());
        return false;
    }
}