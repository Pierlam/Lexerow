using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ProgRun;
public class ProgRunVarMgr
{
    public List<ProgRunVar> ListRunVar { get; private set; } = new List<ProgRunVar>();

    public void Add(ProgRunVar progRunVar)
    {
        ListRunVar.Add(progRunVar);
    }

    public ProgRunVar FindVarByName(string varname)
    {
        return  ListRunVar.FirstOrDefault(v => v.NameEquals(varname));
    }

    /// <summary>
    /// Find the last var. Usefull if the value of the var is a var.
    /// exp a=b
    /// b=12
    /// > will return the var b.
    /// </summary>
    /// <param name="varname"></param>
    /// <returns></returns>
    public ProgRunVar FindLastInnerVarByName(string varname)
    {
        ProgRunVar currProgRunVar = null;

        while (true) 
        {
            currProgRunVar = ListRunVar.FirstOrDefault(v => v.NameEquals(varname));
            if (currProgRunVar == null) return null;

            var v= currProgRunVar.Value as InstrObjectName;
            if(v==null)return currProgRunVar;

            // now find the var value
            varname = v.ObjectName;
        }
    }
}
