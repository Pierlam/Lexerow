using Lexerow.Core.System;
using Lexerow.Core.System.ScriptCompile;

namespace Lexerow.Core.ScriptCompile.Parse;

/// <summary>
/// OnExcel instruction compilation builder.
/// It's a complex instruction, need several stages to build and finalize it.
/// </summary>
public class InstrOnExcelBuilder
{
    /// <summary>
    /// OnExcel build, new stage to manage.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="listVar"></param>
    /// <param name="stkInstr"></param>
    /// <param name="instr"></param>
    /// <param name="listInstrToExec"></param>
    /// <param name="isToken"></param>
    /// <returns></returns>
    public static bool OnExcelBuildOngoing(ExecResult execResult, List<InstrObjectName> listVar, CompilStackInstr stkInstr, InstrBase instr, List<InstrBase> listInstrToExec, out bool isToken)
    {
        isToken = false;
        InstrOnExcel instrOnExcel;

        //--case1: current instr is OnExcel and the stack is empty -> start of OnExcel instr
        if (instr.InstrType == InstrType.OnExcel && stkInstr.Count == 0)
        {
            stkInstr.Push(instr);
            isToken = true;
            return true;
        }

        //--case2: current instr is OnExcel and the stack has one instr, should be OnExcel -> end of OnExcel instr
        if (instr.InstrType == InstrType.OnExcel && stkInstr.Count == 1)
        {
            //it is End OnExcel, last token of the instr OnExcel
            instrOnExcel = stkInstr.Peek() as InstrOnExcel;
            if (instrOnExcel.BuildStage == InstrOnExcelBuildStage.EndOnExcel_End)
            {
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.EndOfExcel;

                // save the instr OnExcel into the list of instr to execute
                listInstrToExec.Add(instrOnExcel);
                stkInstr.Pop();
                isToken = true;
                return true;
            }
        }

        // wrong case
        if (instr.InstrType == InstrType.OnExcel && stkInstr.Count > 1)
        {
            // const value not a string, error
            execResult.AddError(ErrorCode.ParserTokenNotExpected, instr.FirstScriptToken());
            return false;
        }

        // the stack is empty
        if (stkInstr.Count == 0)
            return true;

        // the saved instr is OnExcel
        instrOnExcel = stkInstr.Peek() as InstrOnExcel;
        if (instrOnExcel == null)
            return true;

        isToken = true;

        //--case2: OnExcel next stage -> filename expected
        if (instrOnExcel.BuildStage == InstrOnExcelBuildStage.OnExcel)
        {
            // filename expected. Type string const value or varname
            if (instr.InstrType == InstrType.ConstValue)
            {
                // should be a string
                if (ParserUtils.IsValueString((instr as InstrConstValue).ValueBase))
                {
                    // save the string filename
                    instrOnExcel.InstrFiles = instr;
                    instrOnExcel.BuildStage = InstrOnExcelBuildStage.Files;
                    return true;
                }
                // const value not a string, error
                execResult.AddError(ErrorCode.ParserConstStringValueExpected, instr.FirstScriptToken());
                return false;
            }
            // filename is a variable, should be defined before
            if (instr.InstrType == InstrType.ObjectName)
            {
                // save the string filename
                instrOnExcel.InstrFiles = instr;
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.Files;
                return true;
            }
            // const value not a string, error
            execResult.AddError(ErrorCode.ParserConstStringValueExpected, instr.FirstScriptToken());
            return false;
        }

        //--case3: Files: next stage -> OnSheet/ForEach token expected
        if (instrOnExcel.BuildStage == InstrOnExcelBuildStage.Files)
        {
            if (instr.InstrType == InstrType.OnSheet)
            {
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.OnSheet;
                // nothing more to do, just check the stage
                return true;
            }

            // FirstRow token found
            if (instr.InstrType == InstrType.FirstRow)
            {
                // OnSheet token not found, so create the default OnSheet, SheetNum=1, the first one
                instrOnExcel.CreateOnSheet(instr.FirstScriptToken(), 1);
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.FirstRowValue;
                // push the FirstRow instr on the stack
                stkInstr.Push(instr);
                return true;
            }

            // ForEach token found
            if (instr.InstrType == InstrType.ForEach)
            {
                // OnSheet token not found, so create the default OnSheet, SheetNum=1, the first one
                instrOnExcel.CreateOnSheet(instr.FirstScriptToken(), 1);
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.ForEach;
                return true;
            }

            // ForEachRow token found, same as found Row token.
            if (instr.InstrType == InstrType.ForEachRow)
            {
                // OnSheet token not found, so create the default OnSheet, SheetNum=1, the first one
                instrOnExcel.CreateOnSheet(instr.FirstScriptToken(), 1);
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.Row;
                return true;
            }

            execResult.AddError(ErrorCode.ParserOnSheetExpected, instr.FirstScriptToken());
            return false;
        }

        //--case4: OnSheet: next stage -> SheetNum/SheetName/ForEach/FirstRow
        if (instrOnExcel.BuildStage == InstrOnExcelBuildStage.OnSheet)
        {
            if (instr.InstrType == InstrType.End)
            {
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.EndOnExcel_End;
                return true;
            }

            // SheetNum
            if (instr.InstrType == InstrType.ConstValue)
            {
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.SheetNum;

                // update the OnExcel object!!
                // TODO:

                return true;
            }
            // SheetName
            // TODO:

            // ForEach
            // TODO:

            // FirstRow
            // TODO:

            // error
            execResult.AddError(ErrorCode.ParserOnSheetExpected, instr.FirstScriptToken());
            return false;
        }

        //--case4: FirstRowValue: next stage -> ForEach/ForEachRow
        if (instrOnExcel.BuildStage == InstrOnExcelBuildStage.FirstRowValue)
        {
            // ForEach token found
            if (instr.InstrType == InstrType.ForEach)
            {
                // OnSheet token not found, so create the default OnSheet, SheetNum=1, the first one
                instrOnExcel.CreateOnSheet(instr.FirstScriptToken(), 1);
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.ForEach;
                return true;
            }

            // ForEachRow token found, same as found Row token.
            if (instr.InstrType == InstrType.ForEachRow)
            {
                // OnSheet token not found, so create the default OnSheet, SheetNum=1, the first one
                instrOnExcel.CreateOnSheet(instr.FirstScriptToken(), 1);
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.Row;
                return true;
            }

            execResult.AddError(ErrorCode.ParserOnSheetExpected, instr.FirstScriptToken());
            return false;
        }

        //--case5: ForEach: next stage -> Row
        if (instrOnExcel.BuildStage == InstrOnExcelBuildStage.ForEach)
        {
            if (instr.InstrType == InstrType.Row)
            {
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.Row;
                // nothing more to do, just check the stage
                return true;
            }
            execResult.AddError(ErrorCode.ParserOnSheetExpected, instr.FirstScriptToken());
            return false;
        }

        //--case6: Row: next stage -> If
        if (instrOnExcel.BuildStage == InstrOnExcelBuildStage.Row)
        {
            if (instr.InstrType == InstrType.If)
            {
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.If;

                // push the If instr on the stack
                stkInstr.Push(instr);
                return true;
            }
            execResult.AddError(ErrorCode.ParserOnSheetExpected, instr.FirstScriptToken());
            return false;
        }

        //--case7: Next: next stage -> End OnExcel
        if (instrOnExcel.BuildStage == InstrOnExcelBuildStage.RowNext)
        {
            // next IfThen found in the ForEach Row instr
            if (instr.InstrType == InstrType.If)
            {
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.If;

                // push the If instr on the stack
                stkInstr.Push(instr);
                return true;
            }

            //next token can be Next which close the current ForEach Row
            if (instr.InstrType == InstrType.Next)
            {
                instrOnExcel.BuildStage = InstrOnExcelBuildStage.OnSheet;

                // push the If instr on the stack
                //stkInstr.Push(instr);
                return true;
            }

            execResult.AddError(ErrorCode.ParserOnSheetExpected, instr.FirstScriptToken());
            return false;
        }
        execResult.AddError(ErrorCode.ParserOnSheetExpected, instr.FirstScriptToken());
        return false;
    }

    /// <summary>
    /// Build the If Then instr.
    /// </summary>
    /// <param name="execResult"></param>
    /// <param name="stackInstr"></param>
    /// <param name="instrIfThenElse"></param>
    /// <returns></returns>
    public static bool BuildIfThen(ExecResult execResult, CompilStackInstr stackInstr, InstrIfThenElse instrIfThenElse)
    {
        // the stack should contains the OnExcel instr
        if (stackInstr.Count != 1)
        {
            execResult.AddError(ErrorCode.ParserTokenNotExpected, instrIfThenElse.FirstScriptToken());
            return false;
        }

        InstrOnExcel instrOnExcel = stackInstr.Peek() as InstrOnExcel;

        //--case7: If: next -> got back to Row, to add maybe others IfThen
        if (instrOnExcel.BuildStage != InstrOnExcelBuildStage.If)
        {
            execResult.AddError(ErrorCode.ParserTokenNotExpected, instrIfThenElse.FirstScriptToken());
            return false;
        }

        // get the last OnSheet
        InstrOnSheet instrOnSheet = instrOnExcel.CurrOnSheet;
        instrOnSheet.ListInstrForEachRow.Add(instrIfThenElse);

        // go back to the stage ForEach Row  or on Ext instr to close the current ForEach Row
        instrOnExcel.BuildStage = InstrOnExcelBuildStage.RowNext;

        return true;
    }
}