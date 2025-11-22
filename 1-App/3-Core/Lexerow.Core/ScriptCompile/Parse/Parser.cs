using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.ScriptCompile.Parse;

/// <summary>
/// Syntax Analyzer.
/// received tokens from the Lexer and generate a program of instructions, ready to be executed.
/// </summary>
public class Parser
{
    private IActivityLogger _logger;

    /// <summary>
    /// list of defined variables.
    /// A variable can be created and set several times.
    /// </summary>
    private List<InstrObjectName> _listVar = new List<InstrObjectName>();

    public Parser(IActivityLogger activityLogger)
    {
        _logger = activityLogger;
    }

    /// <summary>
    /// process a script, create instructions to execute.
    /// Analyse source code tokens line by line.
    ///
    /// Comparison sep are re-arranged, exp:<,=  to <=
    /// </summary>
    /// <param name="listSourceCodeLineTokens"></param>
    /// <param name="compiledScript"></param>
    /// <returns></returns>
    public bool Process(Result result, List<ScriptLineTokens> listScriptLineTokens, Program program)
    {
        _logger.LogCompilStart(ActivityLogLevel.Important, "Parser.Process", "script lines Num: " + listScriptLineTokens.Count.ToString());
        _listVar.Clear();

        // no token in the source code! -> error or warning?
        if (listScriptLineTokens.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenExpected, string.Empty);
            return false;
        }

        // process, loop on tokens
        bool res = LoopOnTokens(result, _listVar, listScriptLineTokens, program);

        if (res)
        {
            _logger.LogCompilEnd(ActivityLogLevel.Important, "Parser.Process", "Instr count: " + program.ListInstr.Count().ToString());

            // ok, no error
            return true;
        }

        _logger.LogCompilEndError(null, "SyntaxAnalParseryzer.Process", "Error count: " + result.ListError.Count().ToString());
        return false;
    }

    /// <summary>
    /// Loop on script tokens to parse it and produce instructions to execute.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="listScriptLineTokens"></param>
    /// <param name="listInstrToExec"></param>
    /// <returns></returns>
    private bool LoopOnTokens(Result result, List<InstrObjectName> listVar, List<ScriptLineTokens> listScriptLineTokens, Program program)
    {
        bool res;
        bool isToken = false;

        //--init vars on line tokens
        int currLineTokensIndex = 0;
        ScriptLineTokens currLineTokens = null;
        currLineTokens = listScriptLineTokens[0];

        // final list of compiled instructions to execute
        //listInstrToExec = new List<InstrBase>();

        //--init vars on tokens
        int currTokenIndex = -1;
        ScriptToken currToken = null;

        // temporary save of instr
        CompilStackInstr stackInstr = new CompilStackInstr(_logger);
        while (true)
        {
            // goto the next token, if it exists
            currTokenIndex++;

            if (currTokenIndex >= currLineTokens.ListScriptToken.Count)
            {
                _logger.LogCompilOnGoing(ActivityLogLevel.Important, "Parser.LoopOnTokens", "End Of line reached, Num: " + currLineTokensIndex.ToString());

                // no more token in the current line tokens, process items saved in the stack
                res = ParserStackContentProcessor.ScriptEndLineReached(result, listVar, currLineTokensIndex, stackInstr, program);
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
            if (currToken.Value.Equals("FirstRow"))
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
                result.AddError(ErrorCode.ParserTokenNotExpected, currToken.Value);
                return false;
            }

            //--is it the SetVar equal char? SetVarDecoder
            res = SetVarDecoder.ProcessSetVarEqualChar(result, listVar, stackInstr, currToken, program.ListInstr, out isToken);
            if (!res) break;
            if (isToken) continue;

            //--Is it comparison separator =, >, <, ...?
            if (ParserUtils.IsComparisonSeparator(currToken))
            {
                // process the content of the stack until the If instr
                res = IfPartDecoder.ProcessStackBeforeTokenSepEqualAfterTokenIf(result, listVar, stackInstr, currToken);
                if (!res) break;
                continue;
            }

            //--Is it the Then token?
            res = IfPartDecoder.ProcessStackBeforeTokenThen(result, listVar, stackInstr, currToken, out isToken);
            if (!res) break;
            if (isToken) continue;

            //--is the token the char ) ?  pop the stack until found ( et traite l'expression. parametre d'une fonction/méthode.
            res = TokenCloseBracketProcessor.Do(result, listVar, stackInstr, currToken, program.ListInstr, out bool isListOfParams, out bool isMathExpr, out List<InstrBase> listItem);
            if (!res) break;
            if (isListOfParams)
            {
                // process the fct call, check and set parameters, error saved
                res = FunctionCallParamsProcessor.ProcessFunctionCallParams(_logger, result, listVar, stackInstr, currToken, program.ListInstr, listItem);
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
            res = InstrBuilder.Build(result, currToken, out InstrBase instr);
            if (!res) break;

            // is it the OnExcel instr build ongoing?
            res = InstrOnExcelBuilder.OnExcelBuildOngoing(result, listVar, stackInstr, instr, program.ListInstr, out isToken);
            if (!res) break;
            if (isToken) continue;

            // process special cases: all token of OnExcel instr inline for exp
            res = ProcessSpecialCases(result, listVar, currLineTokensIndex, stackInstr, instr, program, out isToken);
            if (!res) break;
            if (isToken) continue;

            // push it on the stack
            stackInstr.Push(instr);
        }

        // end of tokens parsing, check for errors
        if (!CheckEndParsing(result, stackInstr, program.ListInstr))
            return false;

        return result.Res;
    }

    /// <summary>
    /// Process special cases like:
    ///  all token are inline, pb is on token Next:
    ///     OnExcel "data.xlsx" ForEach Row If A.Cell >10 Then A.Cell=10 Next
    ///
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="currLineTokensIndex"></param>
    /// <param name="stackInstr"></param>
    /// <param name="instr"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool ProcessSpecialCases(Result result, List<InstrObjectName> listVar, int currLineTokensIndex, CompilStackInstr stackInstr, InstrBase instr, Program program, out bool isToken)
    {
        isToken = false;

        if (instr.InstrType == InstrType.Next)
        {
            // special case? Next inline: ..Then A.Cell= 12 Next   or  ..Then fct() Next
            bool res = ParserStackContentProcessor.ScriptEndLineReached(result, listVar, currLineTokensIndex, stackInstr, program);
            if (!res) return false;

            // now process the token Next of the OnExcel instr
            res = InstrOnExcelBuilder.OnExcelBuildOngoing(result, listVar, stackInstr, instr, program.ListInstr, out isToken);
            if (!res) return false;
            return true;
        }

        return true;
    }

    /// <summary>
    /// End of tokens parsing, check for errors
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="listInstrToExec"></param>
    /// <returns></returns>
    private static bool CheckEndParsing(Result result, CompilStackInstr stackInstr, List<InstrBase> listInstrToExec)
    {
        if (!result.Res)
        {
            // clear the list of instructions obtained
            listInstrToExec.Clear();
            return false;
        }

        // nothing in the stack, ok!
        if (stackInstr.Count == 0)
            return true;

        // the stack contains items, error have occured
        InstrOnExcel instrOnExcel = stackInstr.Peek() as InstrOnExcel;
        if (instrOnExcel != null)
        {
            if (instrOnExcel.BuildStage == InstrOnExcelBuildStage.OnSheet)
            {
                result.AddError(ErrorCode.ParserTokenExpected, instrOnExcel.FirstScriptToken(), "Maybe 'End OnExcel' is missing");
                return false;
            }
            result.AddError(ErrorCode.ParserTokenExpected, instrOnExcel.FirstScriptToken());
            return false;
        }

        result.AddError(ErrorCode.ParserTokenExpected, stackInstr.Peek().FirstScriptToken());
        return false;
    }
}