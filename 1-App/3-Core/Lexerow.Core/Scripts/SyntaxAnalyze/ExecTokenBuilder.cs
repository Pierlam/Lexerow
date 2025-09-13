using Lexerow.Core.System;
using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts;
public class ExecTokenBuilder
{
    public static bool Do(ScriptToken scriptToken, out ExecTokBase execTokBase)
    {
        //--script token is a name/id
        if (scriptToken.ScriptTokenType == ScriptTokenType.Name)
        {
            // OpenExcel
            if (scriptToken.Value.Equals("OpenExcel", StringComparison.InvariantCultureIgnoreCase))
            {
                execTokBase = new ExecTokOpenExcel(scriptToken);
                return true;
            }

            // OnExcel
            // TODO:

            // OnSheet
            // TODO:

            // ForEach
            // TODO:

            // Row
            // TODO:

            // Col
            // TODO:

            // if
            if (scriptToken.Value.Equals("if",StringComparison.InvariantCultureIgnoreCase))
            {
                //execTokBase = new ExecTokIf(scriptToken);
                execTokBase = null;
                return true;
            }

            // then
            if (scriptToken.Value.Equals("Then", StringComparison.InvariantCultureIgnoreCase))
            {
                //execTokBase = new ExecTokThen(scriptToken);
                execTokBase = null;
                return true;
            }

            // ExcelCol, exp: A
            // TODO:


            // ExcelCellAddress, exp: A1
            // TODO:

            // if it is not a known keyword, it's an object name, can be: a variable or a user defined function 
            execTokBase = new ExecTokObjectName(scriptToken);
            return true;
        }

        //--script token is a string
        if (scriptToken.ScriptTokenType == ScriptTokenType.String)
        {
            execTokBase = new ExecTokConstValue(scriptToken, scriptToken.Value);
            return true;
        }

        //--script token is a int
        if (scriptToken.ScriptTokenType == ScriptTokenType.Integer)
        {
            execTokBase = new ExecTokConstValue(scriptToken, scriptToken.ValueInt);
            return true;
        }

        //--script token is a double
        if (scriptToken.ScriptTokenType == ScriptTokenType.Double)
        {
            execTokBase = new ExecTokConstValue(scriptToken, scriptToken.ValueDouble);
            return true;
        }


        //--script token is a separator, TODO: préciser?
        // TODO:

        execTokBase = null;
        return true;
    }
}
