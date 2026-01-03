using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ScriptCompile.Parse;

/// <summary>
/// Comparison Parser
/// exp: If A.Cell > 12
/// </summary>
public class ComparisonParser
{
    /// <summary>
    /// Parse the comparison operator. Check that it's a not a SetVar.
    /// Can be found in a If instr.
    /// Just pushed on the stack.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool ParseCompOperator(Result result, CompilStackInstr stackInstr, ScriptToken scriptToken, out bool isToken)
    {     
        isToken = false;

        //--not a comparison  operator =, >, <, ...?
        if (!ParserUtils.IsComparisonOperator(scriptToken))
            return true;

        // the stack contains nothing, error!
        if (stackInstr.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // is it the equal separator? can vbe a SetVar instr
        if (ParserUtils.IsEqualOperator(scriptToken))
        {
            // special case, comparison equal is defined only in a if instr
            if(stackInstr.FindInstrFromTop(InstrType.If) == null)
                // no If found, so it's a SetVar instr, not a comparison
                return true;
        }

        // push the separator comparison into the stack
        InstrBase instrCompOperator = InstrBuilder.CreateSepComparison(scriptToken);
        stackInstr.Push(instrCompOperator);
        isToken = true;
        return true;

    }
}
