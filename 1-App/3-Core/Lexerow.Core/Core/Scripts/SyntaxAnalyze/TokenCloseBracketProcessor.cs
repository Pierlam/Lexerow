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
    /// <param name="stkItems"></param>
    /// <param name="scriptToken"></param>
    /// <param name="listInstr"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool Do(ExecResult execResult, List<InstrObjectName> listVar, Stack<InstrBase> stkItems, ScriptToken scriptToken, List<InstrBase> listInstr, out bool isListOfParams, out bool isMathExpr, out List<InstrBase> listItem)
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
        InstrBase item = stkItems.Peek();
        if (item is InstrOpenBracket)
        {
            // no item in the stack? case: ()
            if (stkItems.Count == 0)
                return true;                    

            InstrBase instStackLast = stkItems.Pop();

            // the last item on the stack is an object name? exp: OpenExcel(f)
            if (instStackLast.IsFunctionCall)
            {
                // TODO: tagg that open and closed brackets are presents
                //instStackLast.OpenCloseBracketsArePresent = true;
            }

            return true;
        }

        //--loop on stacked item: item, a string, a number or an instruction and then , comma sep
        while (true)
        {
            // get the last item from the stack
            if (!stkItems.TryPop(out item))
            {
                // no more item in the stack
                // TODO -> error
            }

            // the current stack item should be an item, a string, a number or an instruction
            //if (item.IsTokenVarName() || item.IsTokenExcelColName() || item.IsTokenExcelCellAddress() || item.IsTokenConstValue() || item.IsInstr())
            if(item is InstrObjectName || item is InstrConstValue)
            {
                // save the item in the fct param list
                listItem.Add(item);
            }
            else
            {
                // TODO: error, item not expected or not yet implemented
                throw new Exception("TODO: type not yet implemented");
            }

            // get the last item from the stack
            if (!stkItems.TryPop(out item))
            {
                // no more item in the stack
                // TODO -> error
            }

            // the next can be the open bracket -> no more fct param
            if (item is InstrOpenBracket)
            {
                // stop the scan of fct parameters
                //isToken = true;
                break;
            }

            // the next one should be the comma sep 
            if (item is InstrComma)
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
            if(stkItems.Count == 0)
            {
                // so it's a math expression
                isMathExpr = true;
                return true; 
            }

            // read the item on the top of the stack
            InstrBase instStackLast= stkItems.Peek();

            // the last item on the stack is an object name? exp: OpenExcel(f)
            if(instStackLast.IsFunctionCall)
            {
                isListOfParams= true;
                return true;
            }
            // other cases, it's a math expression
            isMathExpr = true;
            return true;
        }
        return true;

    }

}
