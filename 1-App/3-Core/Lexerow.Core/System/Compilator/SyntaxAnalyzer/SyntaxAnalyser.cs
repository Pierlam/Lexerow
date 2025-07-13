using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Compilator;

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
    public bool Process(List<SourceCodeLineTokens> listLineTokens, out List<InstrBase> listInstr)
    {

        // no token in the source code! -> error ou warning?
        if (listLineTokens.Count == 0)
        {
            listInstr = null;
            //compiledScript.Errors=XXX
            return true;
        }


        // check for wrong tokens: stringWrong and DoubleWrong
        // TODO: ->error, stop

        // process, loop on tokens 
        //ProcessLoopOnTokens(listLineTokens, compiledScript);
        //if (compiledScript.ListError.Count > 0) return false;

        listInstr = null;
        // ok, no error
        return true;
    }

}
