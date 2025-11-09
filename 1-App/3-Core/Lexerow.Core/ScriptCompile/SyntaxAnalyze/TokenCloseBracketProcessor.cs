using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using NPOI.OpenXmlFormats.Spreadsheet;
using Org.BouncyCastle.Utilities.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ScriptCompile.SyntaxAnalyze;
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
    /// <param name="execResult"></param>
    /// <param name="listVar"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="listInstr"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool Do(ExecResult execResult, List<InstrObjectName> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, List<InstrBase> listInstr, out bool isListOfParams, out bool isMathExpr, out List<InstrBase> listItem)
    {
        isListOfParams = false;
        isMathExpr = false;
        listItem = new List<InstrBase>();

        //--is it ) ?
        if (!scriptToken.Value.Equals(")"))
            // no, bye
            return true;

        // no more item in the stack
        // TODO:  -> error

        //--case 1.1: previous token is (  -> fct()
        InstrBase instr = stackInstr.Peek();
        if (instr is InstrOpenBracket)
        {
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

            //-item before openBracket is an math operator? exp: +()
            // TODO: error? to confirm

            // error, item before the open bracket is not expected
            execResult.AddError(ErrorCode.ParserTokenNotExpected, instBeforeOpenBracket.FirstScriptToken());
            return false;
        }

        //--loop on stacked item: item, a string, a number or an instruction and then , comma sep
        while (true)
        {
            // get the last item from the stack
            if(stackInstr.Count == 0)
            {
                // no more item in the stack
                // TODO -> error
                throw new InvalidOperationException("todo");
            }
            instr= stackInstr.Pop();

            // the current stack item should be an item, a string, a number or an instruction
            //if (item.IsTokenVarName() || item.IsTokenExcelColName() || item.IsTokenExcelCellAddress() || item.IsTokenConstValue() || item.IsInstr())
            if(instr is InstrObjectName || instr is InstrConstValue)
            {
                // save the item in the fct param list
                listItem.Add(instr);
            }
            else
            {
                // TODO: error, item not expected or not yet implemented
                throw new Exception("TODO: type not yet implemented");
            }

            // get the last item from the stack
            if (stackInstr.Count == 0)
            {
                // no more item in the stack
                // TODO -> error
                throw new InvalidOperationException("todo");
            }
            instr = stackInstr.Pop();

            // the next can be the open bracket -> no more fct param
            if (instr is InstrOpenBracket)
            {
                // stop the scan of fct parameters
                //isToken = true;
                break;
            }

            // the next one should be the comma sep 
            if (instr is InstrComma)
            {
                // already a list  of params or type not yet defined so ok
                if(!isMathExpr)
                    isListOfParams |= true;
                continue;
            }

            //
            // comma sep expected!
            //string log = string.Empty;
            //if ((item.IsInstr()))
            //{
            //    log = item.InstrBase.ToString();
            //}
            //else
            //    log = item.SourceCodeToken.Value;

            //var err = ErrorSyntaxAnalyzer.CreateError(ErrorCode.BracketClosedCommaSepExpected, item.SourceCodeToken.LineNum, item.SourceCodeToken.ColNum, log);
            //compiledScript.ListError.Add(err);
            return false;
        }

        // only one item found between open-Closed brackets
        if(listItem.Count == 1)
        {
            // no item in the stack? exp (12)
            if(stackInstr.Count == 0)
            {
                // so it's a math expression
                isMathExpr = true;
                return true; 
            }

            // read the item from the stack
            InstrBase instBeforeOpenBracket= stackInstr.Peek();

            // the last item on the stack is an object name? exp: OpenExcel(f)
            if(instBeforeOpenBracket.IsFunctionCall)
            {
                isListOfParams= true;
                return true;
            }

            // the item is an math operator, before the open bracket? exp: *(2)
            if(ParserUtils.IsMathOperator(instBeforeOpenBracket))
            {
                isMathExpr = true;
                return true;
            }

            // error, wrong object name
            execResult.AddError(ErrorCode.ParserTokenNotExpected, instBeforeOpenBracket.FirstScriptToken());
            return false;
        }
        return true;

    }

}
