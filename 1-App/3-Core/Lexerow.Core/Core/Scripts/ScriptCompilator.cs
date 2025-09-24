using Lexerow.Core.Scripts.LexicalAnalyse;
using Lexerow.Core.Scripts.SyntaxAnalyze;
using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using NPOI.OpenXmlFormats.Shared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts;

/// <summary>
/// Compile script coming from a texte file (or a string) into a list of instructions ready to execute.
/// </summary>
public class ScriptCompilator
{

    CoreData _coreData;
    LexicalAnalyzerConfig lexicalAnalyzerConfig = new LexicalAnalyzerConfig();

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="coreData"></param>
    public ScriptCompilator(CoreData coreData)
    {
        _coreData = coreData;
    }

    /// <summary>
    /// Coimpile the script (source code). tis a list of lines.
    /// Generate a list of instructions, ready to be executed.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="script"></param>
    /// <param name="listInstr"></param>
    /// <returns></returns>
    public ExecResult CompileScript(ExecResult execResult, Script script, out List<InstrBase> listInstr)
    {
        // analyse the source code, line by line
        if (!LexicalAnalyzer.Process(execResult, script, out List<ScriptLineTokens> listScriptLineTokens, lexicalAnalyzerConfig))
        {
            listInstr = new List<InstrBase>();
            return execResult;
        }

        // re-arrange comparison separators, gather them, exp: >,= to >=  ...
        //ComparisonSepMgr.ReArrangeAllComparisonSep(listSourceCodeLineTokens);

        SyntaxAnalyser syntaxAnalyser = new SyntaxAnalyser();
        syntaxAnalyser.Process(execResult, listScriptLineTokens, out listInstr);

        return execResult;
    }

}
