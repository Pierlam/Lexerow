using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.InstrDef.FuncCall;

/// <summary>
/// instruction: CopyHeader(fileSrc, fileTgt)
/// Copy an header from a source file (sheet) to a target file.
/// </summary>
public class InstrFuncCallCopyHeader:InstrBase
{
    public InstrFuncCallCopyHeader(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.FuncCopyHeader;
        IsFunctionCall = true;
    }

    /// <summary>
    /// source file to copy the header from.
    /// can be a string, or an excel file object.
    /// or a list of string/excel files, in that case, the header will be copied from the first file in the list.
    /// </summary>
    public InstrBase InstrSourceFile { get; set; }

    /// <summary>
    /// a string, or an excel file object.
    /// </summary>
    public InstrBase InstrTargetFile { get; set; }

    /// <summary>
    /// If provided, the sheet will be created with the name. Otherwise, the default name will be used.
    /// </summary>
    public InstrBase? InstrSourceSheet { get; set; } = null;

    /// <summary>
    /// If provided, the sheet will be created with the name. Otherwise, the default name will be used.
    /// </summary>
    public InstrBase? InstrTargetSheet { get; set; } = null;

    public override string ToString()
    {
        return "CopyHeader()";
    }

}
