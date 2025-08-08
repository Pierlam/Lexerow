using Lexerow.Core.Scripts;
using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using Lexerow.Core.Tests._05_Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ScriptSyntaxAnalyze;


/// <summary>
/// Test script lexical analyzer.
/// </summary>
[TestClass]
public class ScriptSyntaxAnalyzerBasicTests
{
    /// <summary>
    /// file=OpenExcel("data.xslx")
    /// file; = ; OpenExcel; (; "data.xlsx"; )
    /// </summary>
    [TestMethod]
    public void FileEqOneExcelFileNameOk()
    {
        ScriptToken st;
        ScriptLineTokens slt;

        //-build one line of tokens
        slt= new ScriptLineTokens();
        slt.AddTokenName(1, 1, "file");
        slt.AddTokenSeparator(1, 1, "=");
        slt.AddTokenName(1, 1, "OpenExcel");
        slt.AddTokenSeparator(1, 1, "(");
        slt.AddTokenString(1, 1, "\"data.xlsx\"");
        slt.AddTokenSeparator(1, 1, ")");

        //-build source code lines of tokens
        List<ScriptLineTokens> lt = [slt];
        
        SyntaxAnalyser sa = new SyntaxAnalyser();

        ExecResult execResult = new ExecResult();
        sa.Process(execResult, lt, out List<InstrBase> listInstr);

        //ici();  //to dev!!  sort une liste d'ExeTok
        
        Assert.Fail("todo");
    }

    // test empty!

    // test wrong OpenExcel instr

}
