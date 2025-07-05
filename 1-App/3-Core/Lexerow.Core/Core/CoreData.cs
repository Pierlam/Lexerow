using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core;

public class CoreData
{
    /// <summary>
    /// Possible to create instructions only in build stage.
    /// TODO: move to program
    /// </summary>
    //public CoreStage Stage { get; set; } = CoreStage.Build;

    /// <summary>
    /// TODO: move to Program
    /// </summary>
    //public List<InstrBase> ListInstr { get; set; } = new List<InstrBase>();

    /// <summary>
    /// List of managed programs.
    /// </summary>
    public List<ProgramInstr> ListProgram { get; set; }=new List<ProgramInstr>();

    public ProgramInstr CurrProgramInstr { get; set; } = null;

    public ProgramInstr GetProgramByName(string name)
    {
        if (string.IsNullOrEmpty(name))
            return null;

        foreach (ProgramInstr prog in ListProgram)
        {
            if(prog.Name.Equals(name,StringComparison.InvariantCultureIgnoreCase)) return prog; 
        }

        return null;
    }
}
