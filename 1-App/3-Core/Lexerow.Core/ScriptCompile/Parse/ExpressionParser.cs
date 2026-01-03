using Lexerow.Core.System;
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
/// Parse expression found between mainly If and Then, and also calculation expression.
///  -fct call parameters, exp: fct(a,b,c)
///  -bool expression, exp: (a and b)
///  -comparison expr, exp: (a>b)
///  -calc expr:, exp:      (a+2)
/// </summary>
public class ExpressionParser
{
    public static bool Process(Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, InstrType instrTypeStart)
    {
        // get instr until the start instr: If or (
        int nstrUntilStartCount = stackInstr.GetDistanceFromTop(instrTypeStart) - 1;

        // only one instr on the stack ? exp: If a Then,  (a)

        //--before, on stack, is it a bool value?

        //--before, on stack, is it a fct call? returing a bool value

        //--before, on stack, is it a comparison expr? exp: A.Cell>10
        //ComparisonParser.Parse()


        //--before, on stack, is it a bool expression?  


    }
}
