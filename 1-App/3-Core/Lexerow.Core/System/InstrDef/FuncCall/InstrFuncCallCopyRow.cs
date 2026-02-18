using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.InstrDef.FuncCall;

/// <summary>
/// Main case: CopyRow(fileTarget)
/// later: 
/// 
///   CopyRow(fileTarget, A.Cell, B.Cell, D.Cell)
///   CopyRow(fileTarget, $file.Name, A.Cell, B.Cell, D.Cell)
///   CopyRowToSheet(fileTarget, sheetNum/SheetName, <-values-... )
/// </summary>
public class InstrFuncCallCopyRow : InstrBase
{
    public InstrFuncCallCopyRow(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.FuncCopyRow;
        IsFunctionCall = true;
    }

    public InstrBase InstrTargetFile { get; set; }  

    public override string ToString()
    {
        return "CopyRow()";
    }
}
