using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Core.Scripts;
internal class SyntaxAnalyserUtils
{
    public static bool IsMathOperator(InstrBase instr)
    {
        if (instr is InstrCharPlus) return true;
        if (instr is InstrCharMinus) return true;
        if (instr is InstrCharMul) return true;
        if (instr is InstrCharDiv) return true;

        return false;
    }
}
