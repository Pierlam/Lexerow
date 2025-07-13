using Lexerow.Core.Scripts.Compilator;
using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using NPOI.OpenXmlFormats.Shared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts;

public class ScriptCompilator
{

    CoreData _coreData;
    LexicalAnalyzerConfig lexicalAnalyzerConfig = new LexicalAnalyzerConfig();

    public ScriptCompilator(CoreData coreData)
    {
        _coreData = coreData;
    }

    public ExecResult CompileScript(ExecResult execResult, SourceScript sourceScript, out List<InstrBase> listInstr)
    {
        // analyse the source code, line by line
        LexicalAnalyzer.Process(sourceScript, out List<SourceCodeLineTokens> listSourceCodeLineTokens, lexicalAnalyzerConfig);

        // re-arrange comparison separators, gather them, exp: >,= to >=  ...
        //ComparisonSepMgr.ReArrangeAllComparisonSep(listSourceCodeLineTokens);

        SyntaxAnalyser syntaxAnalyser = new SyntaxAnalyser();
        syntaxAnalyser.Process(listSourceCodeLineTokens, out listInstr);

        // TODO: DEV:
        return new ExecResult();
    }

}
