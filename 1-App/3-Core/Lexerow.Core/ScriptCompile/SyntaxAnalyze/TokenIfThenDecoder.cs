using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ScriptCompile.SyntaxAnalyze;
internal class TokenIfThenDecoder
{
    /// <summary>
    /// Process the content of the stack, because a comparison script token separator have been found.
    /// case1: If A.Cell >
    /// case2: If varBool  NO! when then token is found 
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="listVar"></param>
    /// <param name="stkItems"></param>
    /// <returns></returns>
    public static bool ProcessStackBeforeTokenSepEqualAfterTokenIf(ExecResult execResult, List<InstrObjectName> listVar, CompilStackInstr stkInstr, ScriptToken scriptTokenSepComp)
    {
        if(stkInstr.Count== 0)
        {
            execResult.AddError(ErrorCode.ParserTokenNotExpected, scriptTokenSepComp);
            return false;
        }

        //--is the stack contains A.Cell expression?
        bool res = ParserUtils.ProcessInstrColCellFunc(execResult, stkInstr, scriptTokenSepComp, out bool isInstr);
        if (!res) return false;
        if (isInstr)
        {
            // push the equal sep into the stack
            InstrBase instrSepEqual= InstrBuilder.CreateSepComparison(scriptTokenSepComp);
            stkInstr.Push(instrSepEqual);
            return true;
        }

        //--is it a function call?  exp: If fct()=12
        // TODO:

        // TODO:
        throw new NotImplementedException("ProcessIfPart");    
    }

    /// <summary>
    /// Is if the Then token?
    /// The stack should contains, several cases: 
    /// -first part of the If condition. exp: If, A.Cell, >, 12
    /// -a fct call returning a bool
    /// -a bool variable
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="listVar"></param>
    /// <param name="stkInstr"></param>
    /// <param name="scriptTokenSepComp"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool ProcessStackBeforeTokenThen(ExecResult execResult, List<InstrObjectName> listVar, CompilStackInstr stkInstr, ScriptToken scriptToken, out bool isToken)
    {
        isToken = false;

        // not the token Then? bye
        if (!scriptToken.Value.Equals("Then", StringComparison.InvariantCultureIgnoreCase))
            return true;

        isToken = true;

        InstrThen instrThen = new InstrThen(scriptToken); 

        List<InstrBase> list = new List<InstrBase>();

        InstrIf instrIf = null;

        // extract instr from the stack and put them into a list
        while(true)
        {
            if (stkInstr.Count==0)
            {
                execResult.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
                return false;
            }
            // is it the If Instr?
            InstrBase instr = stkInstr.Peek();
            if (instr.InstrType == InstrType.If)
            {
                instrIf = instr as InstrIf;
                break;
            }

            instr =stkInstr.Pop();
            list.Add(instr);
        }

        // nothing between If and then
        if (list.Count == 0)
        {
            execResult.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // 3 instr between If and Then -> If leftOperand operator rightOperand Then
        if (list.Count==3)
        {
            // check the operator
            InstrSepComparison instrOperator = list[1] as InstrSepComparison;
            if(instrOperator==null)
            {
                execResult.AddError(ErrorCode.ParserSepComparatorExpected, list[1].FirstScriptToken());
                return false;
            }

            // if operator is neither equal nor diff, type of operand can be int, double, NOT string
            // TODO: if one operand is <col>.Cell, all is good
            InstrComparison instrComparison = new InstrComparison(list[2].FirstScriptToken());
            instrComparison.OperandLeft = list[2];
            instrComparison.OperandRight = list[0];
            instrComparison.Operator = instrOperator;

            instrIf.InstrBase= instrComparison;

            // push the token Then on the stack
            stkInstr.Push(instrThen);
            return true;
        }

        // 1 instr between If and Then -> If instr Then
        if (list.Count == 1)
        {
            // check the return type, should be a bool value
            if (list[0].ReturnType != InstrFunctionReturnType.ValueBool)
            {
                execResult.AddError(ErrorCode.ParserReturnTypeWrong, scriptToken);
                return false;
            }

            instrIf.InstrBase = list[0];

            // push the token Then on the stack
            stkInstr.Push(instrThen);
            return true;
        }

        // wrong instr count between If and Then
        execResult.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
        return false;
    }
}
