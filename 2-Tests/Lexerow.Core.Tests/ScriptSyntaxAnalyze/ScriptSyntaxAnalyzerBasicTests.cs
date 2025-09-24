using Lexerow.Core.Scripts.SyntaxAnalyze;
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
    /// file; =; OpenExcel; (; "data.xlsx"; )
    /// 
    /// The compilation return:
    ///  -SetVar:
    ///      InstrRight: ObjectName: file
    ///      InstrLeft:  OpenExcel, p="data.xlsx" 
    /// </summary>
    [TestMethod]
    public void FileEqOpenExcelConstOk()
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
        bool res= sa.Process(execResult, lt, out List<InstrBase> listInstr);

        Assert.IsTrue(res);
        Assert.AreEqual(1, listInstr.Count);

        // SetVar
        Assert.AreEqual(InstrType.SetVar, listInstr[0].InstrType);
        InstrSetVar instrSetVar = listInstr[0] as InstrSetVar;

        // InstrLeft: ObjectName
        InstrObjectName instrObjectName = instrSetVar.InstrLeft as InstrObjectName;
        Assert.IsNotNull(instrObjectName);
        Assert.AreEqual("file", instrObjectName.ObjectName);

        // InstrRight: OpenExcel
        InstrOpenExcel instrOpenExcel = instrSetVar.InstrRight as InstrOpenExcel;
        Assert.IsNotNull(instrOpenExcel);

        // OpenExcel Param
        InstrConstValue instrConstValue = instrOpenExcel.Param as InstrConstValue;
        Assert.IsNotNull(instrConstValue);
        Assert.AreEqual("\"data.xlsx\"", instrConstValue.RawValue);
        Assert.AreEqual("\"data.xlsx\"", (instrConstValue.ValueBase as ValueString).Val);
    }

    /// <summary>
    /// file=OpenOuxcel()		
    /// Error, Wrong/Not expected function name.
    /// </summary>
    [TestMethod]
    public void FileEqOpenOuxcelWrong()
    {
        ici();
    }


	//"john"=OpenExcel(..)    type incorrect, variable attendue
	//12=OpenExcel(...)       type incorrect, variable attendue
 //   file = OpenExcel()        paramètre attendu

 //   file=OpenExcel parenthèses de fonction attendues
 //   file = OpenExcel(f)      variable f n'est pas déclaré avant
	//file=blabla()           Nom fonction inconnue
 //   file = OpenExcel("file.xlsx") load Instruction 'load' après fonction non-autorisé
 //   then = OpenExcel("dd")   nom variable non-authorisé, mot-clé réservé

}
