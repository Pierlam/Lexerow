using Lexerow.Core.System;

namespace Lexerow.Core.ScriptCompile.Parse;

public class ExcelFunctionNameDecoder
{
    public static bool Do(string excelFuncName, out InstrBase excelFunc)
    {
        excelFunc = null;
        if (string.IsNullOrEmpty(excelFuncName)) return false;

        excelFuncName = excelFuncName.Trim();

        if (excelFuncName.Equals("Cell", StringComparison.CurrentCultureIgnoreCase))
        {
            //excelFunc = new ExecTokExcelFuncCell();
            return true;
        }

        //if (excelFuncName.Equals("BgColor", StringComparison.CurrentCultureIgnoreCase))
        //{
        //    excelFunc = new ExecTokExcelFuncBgColor();
        //    return true;
        //}

        //if (excelFuncName.Equals("FgColor", StringComparison.CurrentCultureIgnoreCase))
        //{
        //    excelFunc = new ExecTokExcelFuncFgColor();
        //    return true;
        //}

        return false;
    }
}