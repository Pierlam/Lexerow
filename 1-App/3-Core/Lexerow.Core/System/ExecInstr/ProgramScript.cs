using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Program script: script compiled from a source, txt file or lines.
/// Contains a list of instructions, ready to be executed.
/// </summary>
public class ProgramScript
{
    public ProgramScript(Script script, List<InstrBase> listInstr)
    {
        Script =script;
        ListInstr = listInstr;
    }

    /// <summary>
    /// Possible to create instructions only in build stage.
    /// </summary>
    public CoreStage Stage { get; set; } = CoreStage.Build;

    /// <summary>
    /// list of instruction to execute. Obtains after Lexical analyse and Syntax analyse.
    /// </summary>
    public List<InstrBase> ListInstr { get; set; } = new List<InstrBase>();

    /// <summary>
    /// Script/source code text lines.
    /// From a text file or from a string stream.
    /// </summary>
    public Script Script { get; private set; }
}
