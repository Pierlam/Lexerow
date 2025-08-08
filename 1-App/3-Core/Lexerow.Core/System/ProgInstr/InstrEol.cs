using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public class InstrEol :InstrBase
{
    public InstrEol()
    {
        InstrType = InstrType.Eol;
    }


    public override string ToString()
    {
        return "Type: Eol";
    }

}
