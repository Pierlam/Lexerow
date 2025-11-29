using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.ScriptCompile.Parse;

/// <summary>
/// Manage If part. Its a comparison, several cases.
/// If A.Cell > value
/// If a=12
/// If A.Cell=blank/null
/// </summary>
internal class IfPartDecoder
{
    /// <summary>
    /// Process the content of the stack, because a comparison script token separator have been found.
    /// case1: If A.Cell >
    /// case2: If varBool  NO! when then token is found
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="stkItems"></param>
    /// <returns></returns>
    public static bool ProcessStackBeforeTokenSepEqualAfterTokenIf(Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptTokenSepComp)
    {
        if (stackInstr.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptTokenSepComp);
            return false;
        }

        //--is the stack contains A.Cell expression?
        bool res = ParserUtils.ProcessInstrColCellFunc(result, stackInstr, scriptTokenSepComp, out bool isInstr);
        if (!res) return false;
        if (isInstr)
        {
            // push the equal sep into the stack
            InstrBase instrSepEqual = InstrBuilder.CreateSepComparison(scriptTokenSepComp);
            stackInstr.Push(instrSepEqual);
            return true;
        }

        //--is it a function call?  exp: If fct()=12
        result.AddError(ErrorCode.ParserCaseNotManaged, scriptTokenSepComp, "If fctCall...");
        return false;
    }

    /// <summary>
    /// Is it the Then token?
    /// The stack can contains several cases:
    /// -first part of the If condition. exp: If, A.Cell, >, 12
    /// -a fct call returning a bool
    /// -a bool variable
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptTokenSepComp"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool ProcessStackBeforeTokenThen(Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, out bool isToken)
    {
        isToken = false;

        // not the token Then? bye
        if (!scriptToken.Value.Equals(CoreInstr.InstrThen, StringComparison.InvariantCultureIgnoreCase))
            return true;

        isToken = true;
        InstrThen instrThen = new InstrThen(scriptToken);

        //--case operandLeft operator B.Cell Then ? exp: If A.Cell>B.Cell
        bool res = ParserUtils.ProcessInstrColCellFunc(result, stackInstr, scriptToken, out bool isInstr);
        if (!res) return false;

        if (!MoveInstrToListUntilReachIf(result, stackInstr, scriptToken, out InstrIf instrIf, out List<InstrBase> listInstr))
            return false;

        // nothing between If and then
        if (listInstr.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        //--3 instr between If and Then -> If leftOperand operator rightOperand Then
        if (listInstr.Count == 3)
        {
            // build the 3 items comparison instructions: operandLeft operator operandRight
            if (!BuildInstrComparison(result, listInstr[2], listInstr[1], listInstr[0], out InstrComparison instrComparison))
                return false;

            // check the comparison, if an error occurs, continue the execution!
            CheckInstrComparison(result, instrComparison);

            instrIf.InstrBase = instrComparison;

            // push the token Then on the stack
            stackInstr.Push(instrThen);
            return true;
        }

        //--1 instr between If and Then -> If instr Then
        if (listInstr.Count == 1)
        {
            // check the return type, should be a bool value
            if (listInstr[0].ReturnType != InstrFunctionReturnType.ValueBool)
            {
                result.AddError(ErrorCode.ParserReturnTypeWrong, scriptToken);
                return false;
            }

            instrIf.InstrBase = listInstr[0];

            // push the token Then on the stack
            stackInstr.Push(instrThen);
            return true;
        }

        // wrong instr count between If and Then
        result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
        return false;
    }

    /// <summary>
    /// Build the 3 items comparison instructions: operandLeft operator operandRight.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="instrLeft"></param>
    /// <param name="instrSepComp"></param>
    /// <param name="InstrRight"></param>
    /// <param name="instrComparison"></param>
    /// <returns></returns>
    private static bool BuildInstrComparison(Result result, InstrBase instrLeft, InstrBase instrSepComp, InstrBase instrRight, out InstrComparison instrComparison)
    {
        instrComparison = null;

        // check the operator
        InstrSepComparison instrOperator = instrSepComp as InstrSepComparison;
        if (instrOperator == null)
        {
            result.AddError(ErrorCode.ParserSepComparatorExpected, instrSepComp.FirstScriptToken());
            return false;
        }

        // if operator is neither equal nor diff, type of operand can be int, double, NOT string
        // TODO: if one operand is <col>.Cell, all is good
        instrComparison = new InstrComparison(instrSepComp.FirstScriptToken());
        instrComparison.OperandLeft = instrLeft;
        instrComparison.OperandRight = instrRight;
        instrComparison.Operator = instrOperator;
        return true;
    }

    /// <summary>
    /// Check the comparison instr.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="instrComparison"></param>
    /// <returns></returns>
    private static bool CheckInstrComparison(Result result, InstrComparison instrComparison)
    {
        if (instrComparison.OperandRight.InstrType == InstrType.InstrBlank || instrComparison.OperandLeft.InstrType == InstrType.InstrBlank ||
            instrComparison.OperandRight.InstrType == InstrType.InstrNull || instrComparison.OperandLeft.InstrType == InstrType.InstrNull)
        {
            if (instrComparison.Operator.Operator != SepComparisonOperator.Equal && instrComparison.Operator.Operator != SepComparisonOperator.Different)
            {
                result.AddError(ErrorCode.ParserSepComparatorWrong, instrComparison.FirstScriptToken());
                return false;
            }
        }
        return true;
    }

    /// <summary>
    ///
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="instrIf"></param>
    /// <param name="listInstr"></param>
    /// <returns></returns>
    private static bool MoveInstrToListUntilReachIf(Result result, CompilStackInstr stackInstr, ScriptToken scriptToken, out InstrIf instrIf, out List<InstrBase> listInstr)
    {
        // to save instr between If and Then, in the reverse direction
        listInstr = new List<InstrBase>();

        instrIf = null;

        // extract instr from the stack and put them into a list
        while (true)
        {
            if (stackInstr.Count == 0)
            {
                result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
                return false;
            }
            // is it the If Instr?
            InstrBase instr = stackInstr.Peek();
            if (instr.InstrType == InstrType.If)
            {
                instrIf = instr as InstrIf;
                break;
            }

            instr = stackInstr.Pop();
            listInstr.Add(instr);
        }
        return true;
    }
}