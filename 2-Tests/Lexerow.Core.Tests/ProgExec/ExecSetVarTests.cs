using Lexerow.Core.ExcelLayer;
using Lexerow.Core.ProgExec;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.Tests._05_Common;
using Lexerow.Core.Tests.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.ProgExec;

/// <summary>
/// Test execution of the instruction SetVar.
/// </summary>
[TestClass]
public class ExecSetVarTests :BaseTests
{
    /// <summary>
    /// a= 12
    /// </summary>
    [TestMethod]
    public void ExecSetVarIntValueOk()
    {
        Program program = TestInstrBuilder.CreateProgram();

        //--SetVar #1:   a= 12
        InstrSetVar instrSetVar = TestInstrBuilder.CreateInstrSetVarNameValueInt("a", 12);
        program.ListInstr.Add(instrSetVar);

        //--create the program runner
        ProgramExecutor programExec = new ProgramExecutor(new ActivityLogger(), new ExcelProcessorNpoi());
        Result result = new Result();
        bool res = programExec.Exec(result, program);
        Assert.IsTrue(res);
    }

    // b=a   so b=12
}
