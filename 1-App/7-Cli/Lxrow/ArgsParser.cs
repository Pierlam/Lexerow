using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lxrow;

public class ArgsParser
{
    public static bool Parse(string[] args, out ProgParams progParams, out string errMsg)
    {
        progParams = new ProgParams();

        if (args.Length != 2)
        {
            errMsg = "Error, 2 arguments expected.";
            progParams = null;
            return false;
        }

        // -script
        if (args[0].ToLower() != "-script")
        {
            errMsg = "Error, first argument should be -script.";
            progParams = null;
            return false;
        }

        // script filename to analyze
        if (args[1].Trim().Length == 0)
        {
            errMsg = "Error, filename to analyze is empty.";
            progParams = null;
            return false;
        }
        progParams.InputScriptFile = args[1].Trim();

        // -out
        //if (args[2].ToLower() != "-out")
        //{
        //    errMsg = "Error, Third argument should be -out.";
        //    progParams = null;
        //    return false;
        //}

        //// excel filename to analyze
        //if (args[3].Trim().Length == 0)
        //{
        //    errMsg = "Error, result filename is empty.";
        //    progParams = null;
        //    return false;
        //}
        //progParams.OutputResultFile = args[3].Trim();
        errMsg = string.Empty;
        return true;
    }
}
