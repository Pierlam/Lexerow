using DocumentFormat.OpenXml.Bibliography;
using Lexerow.Core.System;
using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.GenDef;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.InstrDef.FuncCall;
using Lexerow.Core.System.InstrDef.Object;
using Lexerow.Core.Utils;
using OpenExcelSdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.InstrProgExec.ExecFunc;

/// <summary>
/// Instr CreateExcel executor.
/// e.g.:
/// 1/ basic case: CreateExcel("data.xlsx")  
/// 2/ CreateExcel("data.xlsx", "sheet1")  .
/// 3/ with var: CreateExcel(filename)
/// </summary>
public class InstrFuncCreateExcelExecutor
{
    private IActivityLogger _logger;
    private ExcelProcessor _excelProcessor;

    public InstrFuncCreateExcelExecutor(IActivityLogger activityLogger, ExcelProcessor excelProcessor)
    {
        _logger = activityLogger;
        _excelProcessor= excelProcessor;
    }

    public bool ExecFuncCreateExcel(Result result, ProgExecContext ctx, ProgExecVarMgr progExecVarMgr, InstrFuncCallCreateExcel instrCreateExcel)
    {
        _logger.LogExecStart(ActivityLogLevel.Debug, "InstrFuncCreateExcelExecutor.ExecFuncCreateExcel", string.Empty);

        // filename instr need to be executed ?
        if (InstrUtils.NeedToBeExecuted(instrCreateExcel.InstrFileName))
        {
            ctx.StackInstr.Push(instrCreateExcel.InstrFileName);
            return true;
        }

        // sheetname instr need to be executed ?
        if (InstrUtils.NeedToBeExecuted(instrCreateExcel.InstrSheetName))
        {
            ctx.StackInstr.Push(instrCreateExcel.InstrSheetName);
            return true;
        }

        // get the filename
        if (!InstrUtils.GetStringFromInstr(result, progExecVarMgr, instrCreateExcel.InstrFileName, out bool _, out string filename))
            return false;

        // get the sheetname
        string sheetname = CoreDefinitions.DefaultExcelSheetName;
        if (instrCreateExcel.InstrSheetName !=null)
        {
            if (!InstrUtils.GetStringFromInstr(result, progExecVarMgr, instrCreateExcel.InstrSheetName, out bool _, out sheetname))
                return false;
        }

        try
        {
            // create the excel file
            ExcelFile excelFile= _excelProcessor.CreateExcelFile(filename, sheetname);

            // TODO: save the excel file object created
            InstrObjectExcelFile instrObjectExcelFile= new InstrObjectExcelFile(instrCreateExcel.FirstScriptToken(), filename, excelFile);

            // remove the instr from the stack
            ctx.StackInstr.Pop();
            ctx.PrevInstrExecuted = instrObjectExcelFile;
            return true;
        }
        catch (Exception ex)
        {
            var error= result.AddNewError(ErrorCode.ExecUnableCreateExcelFile, instrCreateExcel.FirstScriptToken());
            _logger.LogExecEndError(error, "InstrFuncCreateExcelExecutor.ExecFuncCreateExcel", string.Empty);
            return false;
        }

        return false;

    }
}
