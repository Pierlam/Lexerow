using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts.SyntaxAnalyze;

public class SyntaxAnalyser
{
    /// <summary>
    /// list of defined variables.
    /// A variable can be created and set several times.
    /// </summary>
    List<InstrObjectName> _listVar = new List<InstrObjectName>();

    /// <summary>
    /// process list of source code tokens to create instructions to execute.
    /// Analyse source code tokens line by line.
    /// 
    /// Comparison sep are re-arranged, exp:<,=  to <=
    /// </summary>
    /// <param name="listSourceCodeLineTokens"></param>
    /// <param name="compiledScript"></param>
    /// <returns></returns>
    public bool Process(ExecResult execResult, List<ScriptLineTokens> listScriptLineTokens, out List<InstrBase> listInstr)
    {
        _listVar.Clear();

        // no token in the source code! -> error or warning?
        if (listScriptLineTokens.Count == 0)
        {
            listInstr = null;
            execResult.AddError(new ExecResultError(ErrorCode.SyntaxAnalyzerNoToken));
            return false;
        }

        // check for wrong tokens: stringWrong and DoubleWrong
        // TODO: ->error, stop

        // process, loop on tokens 
        bool res= ProcessLoopOnTokens(execResult, _listVar, listScriptLineTokens, out listInstr);

        if(res)
            // ok, no error
            return true;

        return false;
    }

    public bool ProcessLoopOnTokens(ExecResult execResult, List<InstrObjectName> listVar, List<ScriptLineTokens> listScriptLineTokens, out List<InstrBase> listInstrToExec)
    {
        bool res;
        bool isToken = false;

        //--init vars on line tokens
        int currLineTokensIndex = 0;
        ScriptLineTokens currLineTokens = null;
        currLineTokens = listScriptLineTokens[0];

        // final list of compiled instructions to execute
        listInstrToExec = new List<InstrBase>();

        //--init vars on tokens
        int currTokenIndex = -1;
        ScriptToken currToken = null;

        // temporary save of SourceTokens and instr
        Stack<InstrBase> stkItems = new Stack<InstrBase>();

        while (true) 
        {
            // goto the next token, if it exists
            currTokenIndex++;

            if (currTokenIndex >= currLineTokens.ListScriptToken.Count)
            {
                //RaiseEvent("EndOfLineReached, LineIdx:" + currLineTokensIndex.ToString());

                // no more token in the current line tokens, process items saved in the stack
                res = ScriptEndLineProcessor.ScriptEndLineReached(execResult, listVar, currLineTokensIndex, stkItems, listInstrToExec);
                if (!res) break;

                // no more token in the current line tokens, go to the next one
                currLineTokensIndex++;

                // no more line tokens
                if (currLineTokensIndex >= listScriptLineTokens.Count)
                    break;

                // get the tokens of the source code line and analyse it
                currLineTokens = listScriptLineTokens[currLineTokensIndex];

                currTokenIndex = 0;
            }

            // get the fist token of the line
            currToken = currLineTokens.ListScriptToken[currTokenIndex];

            //--is the token a comment?  dont manage it
            if (currToken.ScriptTokenType == ScriptTokenType.Comment)
                continue;

            //--script token is wrong
            if (currToken.ScriptTokenType == ScriptTokenType.Undefined || currToken.ScriptTokenType == ScriptTokenType.StringBadFormed ||
                currToken.ScriptTokenType == ScriptTokenType.WrongNumber)
            {
                execResult.AddError(new ExecResultError(ErrorCode.SyntaxAnalyzerTokenNotExpected, currToken.Value));
                return false;
            }

            //--is it the SetVar equal char? SetVarDecoder
            res = SetVarDecoder.ProcessSetVarEqualChar(execResult, listVar, stkItems, currToken, listInstrToExec, out isToken);
            if (!res) break;
            if (isToken) continue;

            //--is the token equal or a comparison operator ?  SetVar or If/comparison
            //res = TokenSepEqGreatLessProcessor.Do(stkItems, currToken, compiledScript.ListError, out isToken);
            //if (!res) break;
            //if (isToken) continue;

            //--is the token the char ) ?  pop the stack until found ( et traite l'expression. parametre d'une fonction/méthode.
            res= TokenCloseBracketProcessor.Do(execResult, listVar, stkItems, currToken, listInstrToExec, out bool isListOfParams, out bool isMathExpr, out List<InstrBase> listItem);
            if (!res) break;
            if (isListOfParams)
            {
                // process the fct call, check and set parameters, error saved
                FunctionCallParamsProcessor.ProcessFunctionCallParams(execResult, listVar, stkItems, currToken, listInstrToExec, listItem);
                continue;
            }
            if(isMathExpr)
            {
                // TODO: exp: (23+a)
                //ProcessMathExpression(stkItems, listInstr, listItem);
                continue;
            }

            // move the script token into an exec token
            res = InstrBuilder.Do(execResult, currToken, out InstrBase instr);
            if (!res) break;

            // do checks in some cases
            res=InstrChecker.Do(execResult, listVar, stkItems, instr);
            if (!res) break;

            // push it on the stack
            stkItems.Push(instr);
        }

        // finish the process
        if(!execResult.Result)
        {
            // clear the list of instructions obtained
            listInstrToExec.Clear();
        }

        return execResult.Result;
    }
}
