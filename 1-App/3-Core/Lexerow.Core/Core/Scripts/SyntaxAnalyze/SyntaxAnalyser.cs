using Lexerow.Core.Core.Scripts;
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
        bool res = ProcessLoopOnTokens(execResult, _listVar, listScriptLineTokens, out listInstr);

        if (res)
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

        // temporary save of instr
        Stack<InstrBase> stackInstr = new Stack<InstrBase>();

        while (true)
        {
            // goto the next token, if it exists
            currTokenIndex++;

            if (currTokenIndex >= currLineTokens.ListScriptToken.Count)
            {
                //RaiseEvent("EndOfLineReached, LineIdx:" + currLineTokensIndex.ToString());

                // no more token in the current line tokens, process items saved in the stack
                res = StackContentProcessor.ScriptEndLineReached(execResult, listVar, currLineTokensIndex, stackInstr, listInstrToExec);
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

            //XXX-DEBUG:
            if (currToken.Value.Equals("Then"))
            {
                int a = 12;
            }

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
            res = SetVarDecoder.ProcessSetVarEqualChar(execResult, listVar, stackInstr, currToken, listInstrToExec, out isToken);
            if (!res) break;
            if (isToken) continue;

            //--Is it comparison separator =, >, <, ...?
            if(SyntaxAnalyserUtils.IsComparisonSeparator(currToken))
            {
                // process the content of the stack until the If instr
                res= TokenIfThenDecoder.ProcessStackBeforeTokenSepEqualAfterTokenIf(execResult, listVar, stackInstr, currToken);
                if (!res) break;
                continue;
            }

            //--Is if the Then token?
            res = TokenIfThenDecoder.ProcessStackBeforeTokenThen(execResult, listVar, stackInstr, currToken, out isToken);
            if (!res) break;
            if (isToken) continue;

            //--is the token the char ) ?  pop the stack until found ( et traite l'expression. parametre d'une fonction/méthode.
            res = TokenCloseBracketProcessor.Do(execResult, listVar, stackInstr, currToken, listInstrToExec, out bool isListOfParams, out bool isMathExpr, out List<InstrBase> listItem);
            if (!res) break;
            if (isListOfParams)
            {
                // process the fct call, check and set parameters, error saved
                res = FunctionCallParamsProcessor.ProcessFunctionCallParams(execResult, listVar, stackInstr, currToken, listInstrToExec, listItem);
                if (!res) break;
                continue;
            }
            if (isMathExpr)
            {
                // TODO: exp: (23+a)
                //ProcessMathExpression(stkItems, listInstr, listItem);
                continue;
            }

            // move the script token into an exec token
            res = InstrBuilder.Build(execResult, currToken, out InstrBase instr);
            if (!res) break;

            // is it the OnExcel instr build ongoing?
            res = InstrOnExcelBuilder.OnExcelBuildOngoing(execResult, listVar, stackInstr, instr, listInstrToExec, out isToken);
            if (!res) break;
            if (isToken) continue;

            // process special cases: all token of OnExcel instr inline for exp
            res= ProcessSpecialCases(execResult, listVar, currLineTokensIndex, stackInstr, instr, listInstrToExec, out isToken);
            if (!res) break;
            if (isToken) continue;

            // push it on the stack
            stackInstr.Push(instr);
        }

        // finish the process
        if (!execResult.Result)
        {
            // clear the list of instructions obtained
            listInstrToExec.Clear();
        }

        return execResult.Result;
    }

    /// <summary>
    /// Process special cases like:
    ///  all token are inline, pb is on token Next:
    ///     OnExcel "data.xlsx" ForEach Row If A.Cell >10 Then A.Cell=10 Next
    ///
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="listVar"></param>
    /// <param name="currLineTokensIndex"></param>
    /// <param name="stackInstr"></param>
    /// <param name="instr"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool ProcessSpecialCases(ExecResult execResult, List<InstrObjectName> listVar, int currLineTokensIndex, Stack<InstrBase> stackInstr, InstrBase instr, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;

        if (instr.InstrType == InstrType.Next)
        {
            // special case? Next inline: ..Then A.Cell= 12 Next   or  ..Then fct() Next
            bool res = StackContentProcessor.ScriptEndLineReached(execResult, listVar, currLineTokensIndex, stackInstr, listInstrToExec);
            if (!res) return false;

            // now process the token Next of the OnExcel instr
            res = InstrOnExcelBuilder.OnExcelBuildOngoing(execResult, listVar, stackInstr, instr, listInstrToExec, out isToken);
            if (!res) return false;
            return true;
        }

        return true;
    }

}
