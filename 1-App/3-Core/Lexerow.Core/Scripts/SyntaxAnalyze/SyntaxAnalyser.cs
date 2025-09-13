using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts;

public class SyntaxAnalyser
{
    /// <summary>
    /// process list of source code tokens to create instructions to execute.
    /// Analyse source code tokens line by line.
    /// 
    /// Comparison sep are re-arranged, exp:<,=  to <=
    /// </summary>
    /// <param name="listSourceCodeLineTokens"></param>
    /// <param name="compiledScript"></param>
    /// <returns></returns>
    public bool Process(ExecResult execResult, List<ScriptLineTokens> listScriptLineTokens, out List<ExecTokBase> listInstr)
    {

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
        ProcessLoopOnTokens(execResult, listScriptLineTokens, out listInstr);
        //if (compiledScript.ListError.Count > 0) return false;

        listInstr = null;
        // ok, no error
        return true;
    }

    public bool ProcessLoopOnTokens(ExecResult execResult, List<ScriptLineTokens> listScriptLineTokens, out List<ExecTokBase> listExecTok)
    {
        bool res;
        bool isToken = false;

        //--init vars on line tokens
        int currLineTokensIndex = 0;
        ScriptLineTokens currLineTokens = null;
        currLineTokens = listScriptLineTokens[0];
        listExecTok = new List<ExecTokBase>();

        //--init vars on tokens
        int currTokenIndex = -1;
        ScriptToken currToken = null;

        // temporary save of SourceTokens and instr
        Stack<SyntaxAnalyserItem> stkItems = new Stack<SyntaxAnalyserItem>();

        while (true) 
        {
            // goto the next token, if it exists
            currTokenIndex++;

            if (currTokenIndex >= currLineTokens.ListScriptToken.Count)
            {
                //RaiseEvent("EndOfLineReached, LineIdx:" + currLineTokensIndex.ToString());

                // no more token in the current line tokens, process items saved in the stack
                res = ProcessTokenEndLine(currLineTokensIndex, stkItems, listExecTok);
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

            //--do some checks, according to the current stage
            // TODO: check1: si previous=OpenExcel/SetLogFile et si current token diff from ( alors ->erreur

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

            //--is the token the instr if?
            // TODO:

            //--is it the SetVar equal char? SetVarDecoder
            res = SetVarDecoder.ProcessSetVarEqualChar(execResult, stkItems, currToken, listExecTok, out isToken);
            if (!res) break;
            if (isToken) continue;

            //--is the token equal or a comparison operator ?  SetVar or If/comparison
            //res = TokenSepEqGreatLessProcessor.Do(stkItems, currToken, compiledScript.ListError, out isToken);
            //if (!res) break;
            //if (isToken) continue;

            // TODO: process the current token

            // move the script token into an exec token
            ExecTokenBuilder.Do(currToken, out ExecTokBase execTokBase);
            // push it on the stack
            //var item = SyntaxAnalyserItem.CreateToken(execTokBase);
            //stkItems.Push(item);
        }

        // finish the process
        // TODO:

        return true;
    }

    /// <summary>
    /// no more token in the current line tokens, process items saved in the stack.
    /// cases:
    /// a=12
    /// b=a
    /// if a=12
    /// if sheet.Cell(A,1)=12 
    /// </summary>
    /// <param name="stkItems"></param>
    /// <param name="token"></param>
    /// <param name="compiledScript"></param>
    /// <returns></returns>
    bool ProcessTokenEndLine(int sourceCodeLineIndex, Stack<SyntaxAnalyserItem> stkItems, List<ExecTokBase> listExecTok)
    {
        // no more item in the stack
        if (stkItems.Count == 0) return true;

        // TODO: revoir!!
        //--2 items in the stack, is it SetVar instr?
        if (ProcessTokenEndLineSetVar(sourceCodeLineIndex, stkItems, listExecTok))
            return true;

        //--TODO: gérer if a=12

        //--TODO: gérer if a=12 then

        // case unexpected -> error
        // TODO:        
        return false;
    }

    /// <summary>
    /// a=12
    /// b=a
    /// sheet.Cell(A,1)=12 
    /// </summary>
    /// <param name="stkItems"></param>
    /// <param name="compiledScript"></param>
    /// <returns></returns>
    bool ProcessTokenEndLineSetVar(int sourceCodeLineIndex, Stack<SyntaxAnalyserItem> stkItems, List<ExecTokBase> listExecTok)
    {
        // TODO: so? 
        return false;
    }


}
