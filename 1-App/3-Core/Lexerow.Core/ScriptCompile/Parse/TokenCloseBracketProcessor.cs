using Lexerow.Core.System;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;
using System.Collections.Generic;

namespace Lexerow.Core.ScriptCompile.Parse;

internal class TokenCloseBracketProcessor
{
    /// <summary>
    /// Is the script token the close bracket char?
    ///
    /// 2 main cases:
    /// 1/ it's a list of params pushed to a function call.
    /// 2/ It's a calculation expression.
    ///
    /// case1.1:  fct()
    /// case1.2:  fct(p1)
    /// others: fct(p1, p2)  fcvt(p1, p2, p3,...)
    ///
    ///   can be a calculation expression, exp: (2+3)
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="listInstr">list of instr param (fct call) or instr operand (math expr)</param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool Process(Result result, List<InstrObjectName> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, out bool isListOfParams, out bool isMathExpr, out List<InstrBase> listInstr)
    {
        isListOfParams = false;
        isMathExpr = false;
        listInstr = new List<InstrBase>();

        //--is it ) ?
        if (!scriptToken.Value.Equals(")"))
            // no, bye
            return true;

        // no more item in the stack
        if (stackInstr.Count == 0)
        {
            // error, wrong object name
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        //--case 1.1: previous token is (  -> fct()
        InstrBase instr = stackInstr.Peek();
        if (instr is InstrOpenBracket)
        {
            if (!ProcessOpenBracketNoParam(result, stackInstr, out isListOfParams))
                return false;
            return true;
        }

        //--loop on stacked instr
        if (!ProcessInstrAndSep(result, stackInstr, scriptToken, instr, listInstr, out isListOfParams, out isMathExpr))
            return false;

        // no mroe item in the stack? exp (12), (a)
        if (stackInstr.Count == 0)
        {
            // error, wrong object name
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // it's a math expression, bye
        if (isMathExpr) return true;

        // read the item from the stack
        InstrBase instBeforeOpenBracket = stackInstr.Peek();

        // the last item on the stack is an object name? exp: SelectFiles(f), Date(xx)
        if (instBeforeOpenBracket.IsFunctionCall)
        {
            isListOfParams = true;
            return true;
        }

        // the item is an math operator, before the open bracket? exp: *(2)
        if (ParserUtils.IsMathOperator(instBeforeOpenBracket))
        {
            isMathExpr = true;
            return true;
        }

        // error, wrong object name
        result.AddError(ErrorCode.ParserTokenNotExpected, instBeforeOpenBracket.FirstScriptToken());
        return false;

    }

    /// <summary>
    /// fct call or math expr.
    /// 
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="isListOfParams"></param>
    /// <returns></returns>
    public static bool ProcessOpenBracketNoParam(Result result, CompilStackInstr stackInstr, out bool isListOfParams)
    {
        isListOfParams = false;

        // remove the openBracket from the stack
        stackInstr.Pop();

        // no more item in the stack? case: ()
        if (stackInstr.Count == 0)
            return true;

        // read without remove it the item on the top of the stack
        InstrBase instBeforeOpenBracket = stackInstr.Peek();

        //-item before openBracket is an object name? exp: fct()
        if (instBeforeOpenBracket.IsFunctionCall)
        {
            // tag that open and closed brackets are presents
            // TODO: how to manage that open and close bracket are present?
            //instBeforeOpenBracket.OpenCloseBracketsArePresent = true;

            // it's an empty list of params for a fct call
            isListOfParams = true;
            return true;
        }

        //-item before openBracket is a math operator? exp: +()
        // TODO: error? to confirm

        // error, item before the open bracket is not expected
        result.AddError(ErrorCode.ParserTokenNotExpected, instBeforeOpenBracket.FirstScriptToken());
        return false;
    }

    /// <summary>
    /// Process list of fct call param of operand/operator.
    /// case1: fct(a,b,c)
    /// case2: (a+b+c)
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="instr"></param>
    /// <param name="listInstr"></param>
    /// <param name="isListOfParams"></param>
    /// <param name="isMathExpr"></param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    /// <exception cref="Exception"></exception>
    public static bool ProcessInstrAndSep(Result result, CompilStackInstr stackInstr, ScriptToken scriptToken, InstrBase instr, List<InstrBase> listInstr, out bool isListOfParams, out bool isMathExpr)
    {
        isListOfParams = false;
        isMathExpr = false;

        //--loop on stacked item: item, a string, a number or an instruction and then , comma sep
        while (true)
        {
            if (stackInstr.Count == 0)
            {
                result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
                return false;
            }

            // get the last item from the stack
            instr = stackInstr.Pop();

            // the current stack item should be an item, a string, a number or an instruction
            if (instr is InstrObjectName || instr is InstrValue || instr.IsFunctionCall)
            {
                // save the item in the fct param list
                listInstr.Add(instr);
            }
            else
            {
                result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
                return false;
            }

            // get the last item from the stack
            if (stackInstr.Count == 0)
            {
                result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
                return false;
            }
            instr = stackInstr.Pop();

            // the next can be the open bracket -> no more fct param
            if (instr is InstrOpenBracket)
            {
                // stop the scan of fct parameters
                break;
            }

            // the next one should be the comma sep
            if (instr is InstrComma)
            {
                // math expr found so error
                if(isMathExpr)
                {
                    result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
                    return false;
                }

                isListOfParams = true;
                continue;
            }

            // math operator found
            if(instr is InstrCharPlus ||  instr is InstrCharMinus || instr is InstrCharDiv || instr is InstrCharMul)
            {
                // list of params found so error
                if (isListOfParams)
                {
                    result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
                    return false;
                }

                isMathExpr = true;

                // save the item in the fct param list
                listInstr.Add(instr);
                continue;
            }

            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // put instr in the right order
        listInstr.Reverse();
        return true;
    }
}