using Lexerow.Core.Scripts.LexicalAnalyse;
using Lexerow.Core.Scripts.SyntaxAnalyze;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivityLog;
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
    IActivityLogger _logger;
    
    CoreData _coreData;
    LexicalAnalyzerConfig lexicalAnalyzerConfig = new LexicalAnalyzerConfig();

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="coreData"></param>
    public ScriptCompilator(IActivityLogger activityLogger, CoreData coreData)
    {
        _logger = activityLogger;
        _coreData = coreData;
    }

    /// <summary>
    /// Compile the script (source code). this a list of lines.
    /// Generate a list of instructions, ready to be executed.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="script"></param>
    /// <param name="listInstr"></param>
    /// <returns></returns>
    public ExecResult CompileScript(ExecResult execResult, Script script, out List<InstrBase> listInstr)
    {
        _logger.LogCompilStart(ActivityLogLevel.Important, "CompileScript", script.Name);

        // analyse the source code, line by line
        if (!LexicalAnalyzer.Process(_logger, execResult, script, out List<ScriptLineTokens> listScriptLineTokens, lexicalAnalyzerConfig))
        {
            listInstr = new List<InstrBase>();
            return execResult;
        }

        // re-arrange comparison separators, gather them, exp: >,= to >=  ...
        //ComparisonSepMgr.ReArrangeAllComparisonSep(listSourceCodeLineTokens);

        SyntaxAnalyser syntaxAnalyser = new SyntaxAnalyser(_logger);
        bool res= syntaxAnalyser.Process(execResult, listScriptLineTokens, out listInstr);

        // save the list of instructions build by the compilation stage
        // TODO:

        if(res) 
            _logger.LogCompilEnd(ActivityLogLevel.Important, "CompileScript", script.Name);
        else
            _logger.LogCompilEndError(execResult.ListError[0], "CompileScript", script.Name);


        return execResult;
    }

}
