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
    /// List of managed programs.
    /// </summary>
    public List<ProgramScript> ListProgram { get; set; }=new List<ProgramScript>();

    public ProgramScript CurrProgramScript { get; set; } = null;

    public ProgramScript GetProgramByName(string name)
    {
        if (string.IsNullOrEmpty(name))
            return null;

        foreach (ProgramScript prog in ListProgram)
        {
            if(prog.Script.Name.Equals(name,StringComparison.InvariantCultureIgnoreCase)) return prog; 
        }

        return null;
    }
}
