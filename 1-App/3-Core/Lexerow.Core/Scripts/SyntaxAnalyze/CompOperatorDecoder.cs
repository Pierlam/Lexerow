using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts.SyntaxAnalyze;
public class CompOperatorDecoder
{
    public static bool Do(string strCompOperator, out ExecTokCompOperator execTokCompOperator)
    {
        execTokCompOperator = null;

        if (string.IsNullOrEmpty(strCompOperator)) return false;

        if(strCompOperator.Equals("="))
        {
            execTokCompOperator= new ExecTokCompOperator(strCompOperator, ExecTokCompOperatorType.Equal);
            return true;
        }

        if (strCompOperator.Equals("<>") || strCompOperator.Equals("!="))
        {
            execTokCompOperator = new ExecTokCompOperator(strCompOperator, ExecTokCompOperatorType.NotEqual);
            return true;
        }

        if (strCompOperator.Equals(">"))
        {
            execTokCompOperator = new ExecTokCompOperator(strCompOperator, ExecTokCompOperatorType.GreaterThan);
            return true;
        }

        if (strCompOperator.Equals(">="))
        {
            execTokCompOperator = new ExecTokCompOperator(strCompOperator, ExecTokCompOperatorType.GreaterOrEqualThan);
            return true;
        }

        if (strCompOperator.Equals("<"))
        {
            execTokCompOperator = new ExecTokCompOperator(strCompOperator, ExecTokCompOperatorType.LessThan);
            return true;
        }

        if (strCompOperator.Equals("<="))
        {
            execTokCompOperator = new ExecTokCompOperator(strCompOperator, ExecTokCompOperatorType.LessOrEqualThan);
            return true;
        }

        return false;
    }
}
