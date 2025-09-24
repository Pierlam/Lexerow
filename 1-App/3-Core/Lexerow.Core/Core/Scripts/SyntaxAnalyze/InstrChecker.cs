using Lexerow.Core.System;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts.SyntaxAnalyze;
internal class InstrChecker
{
    public static bool Do(ExecResult execResult, List<InstrObjectName> listVar, Stack<InstrBase> stkItems, InstrBase instr)
    {
        if(instr.InstrType == InstrType.OpenExcel)
        {
            // the stack should contains one item

            // last item in the stack should a var

            //         execResult.AddError(new ExecResultError(ErrorCode.SyntaxAnalyzerTokenNotExpected, scriptToken.Value));


            return true;
        }

        return true;
    }
}
