using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
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
    public bool Process(ExecResult execResult, List<ScriptLineTokens> listScriptLineTokens, out List<InstrBase> listInstr)
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

    public bool ProcessLoopOnTokens(ExecResult execResult, List<ScriptLineTokens> listScriptLineTokens, out List<InstrBase> listInstr)
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

        // temporary save of SourceTokens and instr
        Stack<SyntaxAnalyserItem> stkItems = new Stack<SyntaxAnalyserItem>();

        while (true) 
        {
            // goto the next token, if it exists
            currTokenIndex++;

            // use ExcelFunctionNameDecoder

            //DEBUG:
            throw new Exception("SyntaxAnalyzer.ProcessLoopOnTokens: to implement");
        }

    }
}
