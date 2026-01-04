using Lexerow.Core.System;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Utils;

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
    /// Is it the token And or Or?  If operand And
    /// - 2 main cases: 
    ///   it's the first bool operator found, If Operand And
    ///   It's the second or more bool operator foud, exp: If Operand1 And Operand2 And
    ///   
    /// Just check it and push the bool operator on the stack.
    /// The content will be finally parsed when the instr Then will be found.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool ProcessTokenAndOr(Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, out bool isToken)
    {
        isToken = false;

        // not the token And/Or? bye
        if (!scriptToken.Value.Equals(CoreInstr.InstrAnd, StringComparison.InvariantCultureIgnoreCase) && !scriptToken.Value.Equals(CoreInstr.InstrOr, StringComparison.InvariantCultureIgnoreCase))
            return true;

        isToken = true;

        //--case operandLeft operator B.Cell Then ? exp: If/And/Or A.Cell> B.Cell -> parse the right operand
        //bool res = ParserUtils.ProcessInstrColCellFunc(result, stackInstr, scriptToken, out bool isInstr);
        //if (!res) return false;

        // parse l'expression!comme pour Then et )
        //ici;

        // save the bool operator And/Or on the stack
        InstrBase instrBoolOperator=InstrUtils.CreateBoolOperator(scriptToken);
        stackInstr.Push(instrBoolOperator);
        return true;
    }

    /// <summary>
    /// Is it the Then token?
    /// Instructions before the comparison operator is already parsed, exp: A, ., Cell -> A.Cell
    /// The stack can contains several cases:
    /// -first part of the If condition. exp: If, A.Cell, >, 12
    /// -first part of the If condition. exp: If, A.Cell, >, B, ., Cell
    /// -a fct call returning a bool
    /// -a bool variable
    /// 
    /// At the ned of the process, instr if contains the comparison or the bool expr. Stay in the stack.
    /// And the instr Then is pushed on top of the stack
    /// >OnExcel, If, Then
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

        //--case operandLeft operator B.Cell Then ? exp: If A.Cell> B.Cell -> parse the right operand
        bool res = ParserUtils.ProcessInstrColCellFunc(result, stackInstr, scriptToken, out bool isInstr);
        if (!res) return false;

        // extract instr from the stack and put them into a list (so in the right order)
        if (!MoveInstrToListUntilReachIf(result, stackInstr, scriptToken, out InstrIf instrIf, out List<InstrBase> listInstr))
            return false;

        // is it a boolean expression? contains And or Or instr
        if(!ProcessBoolExpr(result, listVar, stackInstr, instrIf, instrThen, listInstr, out bool isTokenExpr)) return false;
        if(isTokenExpr) return true;

        // nothing between If and then
        if (listInstr.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        //--1 instr between If and Then -> If instr Then
        if (listInstr.Count == 1)
        {
            if (!CheckInstrReturnBoolValue(result, listInstr[0]))
                return false;

            instrIf.InstrBase = listInstr[0];

            // push the token Then on the stack
            stackInstr.Push(instrThen);
            return true;
        }

        //--3 instr between If and Then -> If leftOperand operator rightOperand Then
        if (listInstr.Count == 3)
        {
            // build the 3 items comparison instructions: operandLeft operator operandRight
            if (!BuildInstrComparison(result, listInstr[0], listInstr[1], listInstr[2], out InstrComparison instrComparison))
                return false;

            instrIf.InstrBase = instrComparison;

            // push the token Then on the stack
            stackInstr.Push(instrThen);
            return true;
        }

        // wrong instr count between If and Then
        result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
        return false;
    }

    /// <summary>
    /// Is it a boolean expression? contains And or Or instr.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="stackInstr"></param>
    /// <param name="listInstr"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool ProcessBoolExpr(Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, InstrIf instrIf, InstrThen instrThen, List<InstrBase> listInstr, out bool isToken)
    {
        isToken = false;

        // contains some and?
        int andCount= listInstr.Count(x => x.InstrType == InstrType.And);
        int orCount = listInstr.Count(x => x.InstrType == InstrType.Or);

        if (andCount == 0 & orCount == 0)
            // not a bool expr
            return true;

        isToken = true;

        // contains And and Or, not possible
        if(andCount > 0 && orCount > 0)
        {
            // wrong instr count between If and Then
            result.AddError(ErrorCode.ParserBoolExprMixAndOrNotAllowed, listInstr[0].FirstScriptToken());
            return false;
        }

        InstrBoolExprOperator oper= InstrBoolExprOperator.And;
        if (orCount > 0) oper = InstrBoolExprOperator.Or;

        // create a bool expression
        InstrBoolExpr instrBoolExpr = new InstrBoolExpr(listInstr[0].FirstScriptToken(), oper);

        List<InstrBase> listInstrExtract = new List<InstrBase>();

        // build each part of the bool expression
        int i = 0;
        while(i<listInstr.Count)
        {
            if(listInstr[i].InstrType != InstrType.And && listInstr[i].InstrType != InstrType.Or)
            { 
                listInstrExtract.Add(listInstr[i]); 
                i++;
                continue;
            }
            if(!ProcessBoolExprSubPart(result, instrBoolExpr, listInstrExtract, listInstr[0]))
                return false;

            // goto next instr to scan
            i++;
        }

        // finish the job
        if (listInstrExtract.Count > 0)
        {
            if (!ProcessBoolExprSubPart(result, instrBoolExpr, listInstrExtract, listInstr[0]))
                return false;
        }

        // set the bool exp into the If instr
        instrIf.InstrBase = instrBoolExpr;

        // save the Then instr on the stack
        stackInstr.Push(instrThen);
        return true;
    }

    /// <summary>
    /// Process a sub part of instr: If subart1 And subPart2 ...
    /// </summary>
    /// <param name="result"></param>
    /// <param name="instrBoolExpr"></param>
    /// <param name="listInstr"></param>
    /// <returns></returns>
    static bool ProcessBoolExprSubPart(Result result, InstrBoolExpr instrBoolExpr, List<InstrBase> listInstr, InstrBase firstInstr)
    {
        // only one instr? should return a bool value
        if (listInstr.Count == 1)
        {
            if (!CheckInstrReturnBoolValue(result, listInstr[0]))
                return false;

            instrBoolExpr.ListOperand.AddRange(listInstr);
            listInstr.Clear();
            return true;
        }

        // 3 instr? should be a comparison instr
        if (listInstr.Count == 3)
        {
            // build the 3 items comparison instruction: operandLeft operator operandRight
            if (!BuildInstrComparison(result, listInstr[0], listInstr[1], listInstr[2], out InstrComparison instrComparison))
                return false;

            instrBoolExpr.ListOperand.Add(instrComparison);
            listInstr.Clear();
            return true;
        }
        
        ScriptToken scriptToken = firstInstr.FirstScriptToken();
        if (listInstr.Count > 0)
            scriptToken = listInstr[0].FirstScriptToken();

        // wrong instr count between If and Then
        result.AddError(ErrorCode.ParserBoolExprWrong, scriptToken);
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

        // check the comparison, if an error occurs, continue the execution!
        if (!CheckInstrComparison(result, instrComparison))
            return false;

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
    /// move all instr of the stack from the top until reach the instr if.
    /// Reverse the index of items of the list, so items are now in the order they were pushed on the stack.
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

        // reverse the index of items of the list
        listInstr.Reverse();
        return true;
    }

    static bool CheckInstrReturnBoolValue(Result result, InstrBase instr)
    {
        // check the return type, should be a bool value
        if (instr.ReturnType != InstrReturnType.ValueBool)
        {
            result.AddError(ErrorCode.ParserReturnTypeWrong, instr.FirstScriptToken());
            return false;
        }

        return true;    
    }

}