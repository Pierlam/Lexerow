using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Utils;

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
    private List<InstrNameObject> _listVar = new List<InstrNameObject>();

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
        _logger.LogCompilStart(ActivityLogLevel.Info, "Parser.Process", "script lines Num: " + listScriptLineTokens.Count.ToString());
        _listVar.Clear();

        // no token in the source code! -> error or warning?
        if (listScriptLineTokens.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenExpected, string.Empty);
            return false;
        }

        // process tokens, line by line
        bool res = PerformTokensLineByLine(result, _listVar, listScriptLineTokens, program);

        if (res)
        {
            _logger.LogCompilEnd(ActivityLogLevel.Info, "Parser.Process", "Instr count: " + program.ListInstr.Count().ToString());

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
    private bool PerformTokensLineByLine(Result result, List<InstrNameObject> listVar, List<ScriptLineTokens> listScriptLineTokens, Program program)
    {
        bool res;
        bool isToken = false;

        //--init vars on line tokens
        int currLineTokensIndex = 0;
        ScriptLineTokens currLineTokens = null;
        currLineTokens = listScriptLineTokens[0];

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
                _logger.LogCompilOnGoing(ActivityLogLevel.Debug, "Parser.PerformTokensLineByLine", "End Of line reached, Num: " + currLineTokensIndex.ToString() + ", Line: " + currLineTokens.ScriptLine);

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
            if (currToken.Value.Equals(")"))
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

            //--is it the Cell or BgColor or FgColor token?
            if (!InstrColCellFuncParser.Perform(result, stackInstr, currToken, out isToken)) break;
            if (isToken) continue;

            //--is it an operator comparison ? =,<,>,<=, >= (and not a set var isntr)
            if (!ComparisonParser.ProcessCompOperator(result, stackInstr, currToken, out isToken)) break;
            if (isToken) continue;


            //--is it the SetVar equal char?
            if (!SetVarParser.PerformSetVar(result, listVar, stackInstr, currToken, out isToken)) break;
            if (isToken) continue;

            //--Is it the And/or token?
            if (!PerformTokenAndOr(result, listVar, stackInstr, currToken, out isToken)) break;
            if (isToken) continue;

            //--Is it the comma "," token?
            // TODO: process the expr: between the comma and the ) or a another ,
            if (!PerformTokenComma(result, listVar, stackInstr, currToken, out isToken)) break;
            if (isToken) continue;

            //--Is it the Then token?
            if (!PerformTokenThen(result, listVar, stackInstr, currToken, out isToken)) break;
            if (isToken) continue;

            //--is the token the char ) ?  
            if (!PerfomTokenRightBracket(_logger, result, listVar, stackInstr, currToken, program, out isToken)) break;
            if (isToken) continue;


            // move the script token into an exec token
            res = InstrBuilder.Build(result, currToken, out InstrBase instr);
            if (!res) break;

            // is it the OnExcel instr build ongoing?
            res = InstrOnExcelBuilder.OnExcelBuildOngoing(result, listVar, stackInstr, instr, program.ListInstr, out isToken);
            if (!res) break;
            if (isToken) continue;

            // process special cases: all token of OnExcel instr inline for exp
            res = PerformSpecialCases(result, listVar, currLineTokensIndex, stackInstr, instr, program, out isToken);
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
    /// Is it the token And or Or?  If operand And
    /// If .. And, If(...And 
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
    public static bool PerformTokenAndOr(Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, out bool isToken)
    {
        isToken = false;

        // not the token And/Or? bye
        if (!scriptToken.Value.Equals(CoreDefinitions.InstrAnd, StringComparison.InvariantCultureIgnoreCase) && !scriptToken.Value.Equals(CoreDefinitions.InstrOr, StringComparison.InvariantCultureIgnoreCase))
            return true;

        isToken = true;

        // nothing in the stack, error
        if (stackInstr.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // Find the start instr: And, Or, ( or If 
        InstrBase instrAndOrOpenBracketIf = stackInstr.FindFirstInstrFromTop(InstrType.And, InstrType.Or);
        if(instrAndOrOpenBracketIf == null)
        {
            // Find the start instr:  ( or If 
            instrAndOrOpenBracketIf = stackInstr.FindFirstInstrFromTop(InstrType.OpenBracket, InstrType.If);
        }
        if(instrAndOrOpenBracketIf==null)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // parse the expression between If and Then, should return a bool value: bool expr or comparison or bool value or fct retunrg a bool value
        if (!ExpressionParser.Perform(result, listVar, stackInstr, scriptToken, instrAndOrOpenBracketIf.InstrType, out List<InstrBase> listInstrOut))
            return false;

        // only one expr instr expected
        if (listInstrOut.Count != 1)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // save the parsed instr between If/And/Or/( and And/or
        stackInstr.Push(listInstrOut[0]);

        // save the bool operator And/Or on the stack
        InstrBase instrBoolOperator = InstrUtils.CreateBoolOperator(scriptToken);
        stackInstr.Push(instrBoolOperator);
        return true;
    }

    /// <summary>
    /// Process the expr: between the comma and the ( or a another ,
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool PerformTokenComma(Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, out bool isToken)
    {
        isToken = false;

        // not the token ,? bye
        if (!scriptToken.Value.Equals(CoreDefinitions.InstrComma, StringComparison.InvariantCultureIgnoreCase))
            return true;

        isToken = true;

        // nothing in the stack, error
        if (stackInstr.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // parse the sub part
        InstrBase instrStart = stackInstr.FindFirstInstrFromTop(InstrType.OpenBracket, InstrType.Comma);
        if(instrStart==null)
        {
            // errror?
        }

        List<InstrBase> listInstrOut;

        // parse the expression between , and , or (
        if (!ExpressionParser.Perform(result, listVar, stackInstr, scriptToken, instrStart.InstrType, out listInstrOut))
            return false;

        // only one result instr expected  -> NON IN FACT, can be a math expr, a bool expr, a comp expr
        if (listInstrOut.Count != 1)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }
        stackInstr.Push(listInstrOut[0]);

        // push now the comma
        InstrComma instrComma = new InstrComma(scriptToken);
        stackInstr.Push(instrComma);
        return true;
    }

    /// <summary>
    /// Is it the Then token? so process the if condition.
    /// Comment: right bracket placed before are alreay processed, exp: if (a > b) Then
    /// Analyse instr back to If, can be:
    ///  -a bool value
    ///  -comparison expr, exp: a>b
    ///  -bool expression, exp: a and b
    ///  -calc expr, not allowed but possible in inner instr.
    ///   
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="stackInstr"></param>
    /// <param name="token"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool PerformTokenThen(Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, out bool isToken)
    {
        isToken = false;

        // not the token Then? bye
        if (!scriptToken.Value.Equals(CoreDefinitions.InstrThen, StringComparison.InvariantCultureIgnoreCase))
            return true;

        isToken = true;

        // nothing in the stack, error
        if (stackInstr.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        List<InstrBase> listInstrOut;

        // parse the sub part if And/Or exists, exp: If .. And A.Cell>10 Then
        InstrBase instrAndOr = stackInstr.FindFirstInstrFromTop(InstrType.And, InstrType.Or);
        if (instrAndOr != null)
        {
            // parse the expression between If and Then, should return a bool value: bool expr or comparison or bool value or fct retunrg a bool value
            if (!ExpressionParser.Perform(result, listVar, stackInstr, scriptToken, instrAndOr.InstrType, out listInstrOut))
                return false;
            // only one result instr expected
            if (listInstrOut.Count != 1)
            {
                result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
                return false;
            }
            stackInstr.Push(listInstrOut[0]);
        }

        // find the If instr in the stack, must exists
        InstrIf instrIf = stackInstr.FindInstrFromTop(InstrType.If) as InstrIf;
        if (instrIf == null)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // parse the expression between If and Then, should return a bool value: bool expr or comparison or bool value or fct retunrg a bool value
        if (!ExpressionParser.Perform(result, listVar, stackInstr, scriptToken, InstrType.If, out listInstrOut))
            return false;

        // only one expr instr expected
        if(listInstrOut.Count != 1)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        if (listInstrOut[0].InstrType == InstrType.BoolExpr || listInstrOut[0].InstrType == InstrType.Comparison || listInstrOut[0].ReturnType== InstrReturnType.ValueBool) 
        {
            // get If instr and add it
            instrIf.InstrBase = listInstrOut[0];

            // create the Then instr and push it on the stack
            InstrThen instrThen = new InstrThen(scriptToken);
            stackInstr.Push(instrThen);
            return true;    
        }


        result.AddError(ErrorCode.ParserBoolExprWrong, listInstrOut[0].FirstScriptToken());
        return false;
    }

    /// <summary>
    /// Is the token right bracket ) ?  can be
    /// can be: 
    ///  -fct call parameters, exp: fct(a,b,c)
    ///  -bool expression, exp: (a and b)
    ///  -comparison expr, exp: (a>b)
    ///  -calc expr:, exp:      (a+2)
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="stackInstr"></param>
    /// <param name="token"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool PerfomTokenRightBracket(IActivityLogger logger, Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, Program program, out bool isToken)
    {
        isToken = false;

        //--is it ) ?
        if (!scriptToken.Value.Equals(")"))
            // no, bye
            return true;

        isToken = true;

        // nothing in the stack, error
        if (stackInstr.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenExpected, scriptToken);
            return false;
        }

        List<InstrBase> listInstrOut;

        // parse the sub part if And/Or exists, exp: fct(.. And A.Cell>10)
        InstrBase instrAndOr = stackInstr.FindFirstInstrFromTop(InstrType.And, InstrType.Or);
        if (instrAndOr != null)
        {
            if (!ExpressionParser.Perform(result, listVar, stackInstr, scriptToken, instrAndOr.InstrType, out listInstrOut))
                return false;

            // only one result instr expected
            if (listInstrOut.Count != 1)
            {
                result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
                return false;
            }
            stackInstr.Push(listInstrOut[0]);
        }

        // parse the sub part if comma exists, exp: fct(.., A.Cell>10)
        // TODO:

        // find the ( instr in the stack, must exists
        InstrBase instrStart = stackInstr.FindInstrFromTop(InstrType.OpenBracket);
        if (instrStart == null)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // Process the expression 
        if (!ExpressionParser.Perform(result, listVar, stackInstr, scriptToken, instrStart.InstrType, out listInstrOut))
            return false;

        // remove the ( from the stack
        stackInstr.Pop();

        // nothing in the stack, error
        if (stackInstr.Count == 0)
        {
            result.AddError(ErrorCode.ParserTokenNotExpected, scriptToken);
            return false;
        }

        // if <instr>  or again (
        InstrBase instrBase=stackInstr.Peek();
        if(instrBase.InstrType == InstrType.If || instrBase.InstrType == InstrType.OpenBracket)
        {
            stackInstr.Push(listInstrOut[0]);
            return true;
        }

        // process the fct call, check and set parameters, error saved
        return FunctionCallParamsParser.PerformFunctionCallParams(logger, result, listVar, stackInstr, scriptToken, program, listInstrOut);
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
    private static bool PerformSpecialCases(Result result, List<InstrNameObject> listVar, int currLineTokensIndex, CompilStackInstr stackInstr, InstrBase instr, Program program, out bool isToken)
    {
        isToken = false;

        //--instr Next
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

        //--a=-7, curr=7. Stack> Minus; SetVar
        if (!ProcessNegativeValueNumber(result, stackInstr, instr, out isToken))
            return false;
        if (isToken) return true;

        //--If a=-1; Stack IN> ??

        //--Then a=-1; Stack IN> ??

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

    /// <summary>
    /// case1: a=-7,
    /// instr=7, Stack> Minus; SetVar
    ///
    /// case2: A.Cell>10
    /// instr=10, Stack> Minus; SepComparison; ColCellFunc; If; OnExcel
    /// </summary>
    /// <param name="result"></param>
    /// <param name="stackInstr"></param>
    /// <param name="instr"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    private static bool ProcessNegativeValueNumber(Result result, CompilStackInstr stackInstr, InstrBase instr, out bool isToken)
    {
        isToken = false;

        // the current instr is not an int or a double?
        if (!InstrUtils.IsValueInt(instr) && InstrUtils.IsValueDouble(instr)) return true;

        // not enought instr in the stack
        if (stackInstr.Count < 2) return true;

        // read the instr in the top of the stack
        InstrCharMinus instrCharMinus = stackInstr.Peek() as InstrCharMinus;
        if (instrCharMinus == null) return true;

        //--case1: the next instr is the SetVar instr?
        if ((stackInstr.ReadInstrBeforeTop() as InstrSetVar) != null)
        {
            InstrUtils.MergeInstrMinus(instrCharMinus, instr as InstrValue);
            // remove the charMinus
            stackInstr.Pop();

            stackInstr.Push(instr);
            isToken = true;
            return true;
        }

        //--case2: the next instr is the SepComparison instr?
        if ((stackInstr.ReadInstrBeforeTop() as InstrSepComparison) != null)
        {
            InstrUtils.MergeInstrMinus(instrCharMinus, instr as InstrValue);
            // remove the charMinus
            stackInstr.Pop();

            stackInstr.Push(instr);
            isToken = true;
            return true;
        }

        return true;
    }
}