using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core;

public class ExecCompileInstrMgr
{
    /// <summary>
    /// check all instruction, one by one.
    /// </summary>
    /// <param name="listInstr"></param>
    /// <returns></returns>
    public static ExecResult CheckAllInstr(List<InstrBase> listInstr)
    {
        ExecResult execResult = new ExecResult();

        foreach (InstrBase instrBase in listInstr)
        {
            //--instr IfThen?
            InstrIfColThen instrIfColThen = instrBase as InstrIfColThen;
            if (instrIfColThen != null)
            {
                execResult.AddError(new ExecResultError(ErrorCode.InstrNotAllowed, instrIfColThen.InstrType.ToString()));
                continue;
            }

            // TODO: others not allowed!!

            //--instr InstrForEachRowIfThen?
            InstrOnExcelForEachRowIfThen instrForEachRowIfThen = instrBase as InstrOnExcelForEachRowIfThen;
            if (instrForEachRowIfThen != null)
            {
                var er = CheckInstrForEachRowIfThen(instrForEachRowIfThen);
                execResult.AddListError(er.ListError);
            }
        }

        return execResult;
    }

    static ExecResult CheckInstrForEachRowIfThen(InstrOnExcelForEachRowIfThen instrForEachRowIfThen)
    {
        ExecResult execResult = new ExecResult();

        // check instr if 
        if (instrForEachRowIfThen.ListInstrIfColThen.Count == 0)
            execResult.AddError(new ExecResultError(ErrorCode.AtLeastOneInstrIfColThenExpected, null));

        // check instr then
        foreach (InstrIfColThen instrIfColThen in instrForEachRowIfThen.ListInstrIfColThen)
        {
            if (instrIfColThen.ListInstrThen.Count == 0)
                execResult.AddError(new ExecResultError(ErrorCode.AtLeastOneInstrThenExpected, null));
        }

        return execResult;
    }

}
