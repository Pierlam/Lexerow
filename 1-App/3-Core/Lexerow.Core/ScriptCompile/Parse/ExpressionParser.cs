using Lexerow.Core.System;
using Lexerow.Core.System.InstrDef;
using Lexerow.Core.System.ScriptCompile;
using Lexerow.Core.System.ScriptDef;
using Lexerow.Core.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Lexerow.Core.ScriptCompile.Parse;

/// <summary>
/// Parse expression found between mainly If and Then, or between ( and ).
/// The script token is Then or ).
/// The number o instr should be an odd number: 1, 3, 5, ... 
/// 
///  -fct call parameters, exp: fct(a,b,c)
///  -bool expression, exp: a and b,  or: is null , or: is not null
///  -comparison expr, exp: a>b
///  -calc expr:, exp:      a+2
/// </summary>
public class ExpressionParser
{
    /// <summary>
    /// Process the instr between If and Then, or between ( and )
    /// called when: And/Or, Then ) is found.
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listVar"></param>
    /// <param name="stackInstr"></param>
    /// <param name="scriptToken"></param>
    /// <param name="instrTypeStart"></param>
    /// <param name="listInstrBaseOut"></param>
    /// <returns></returns>
    public static bool Perform(Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, InstrType instrTypeStart, out List<InstrBase> listInstrBaseOut)
    {
        listInstrBaseOut = new List<InstrBase>();

        // get instr until the start instr: If or (
        int instrCount = stackInstr.GetDistanceFromTop(instrTypeStart) - 1;

        // nothng between start and end instr, can be a fct call without any parameter, fct()
        if (instrCount == 0)
            return true;

        //--before, on stack, is it a bool value? exp: If a Then, If (a) Then or a fct returning a bool value, fct()
        if (instrCount == 1)
        {
            InstrBase instrBase = stackInstr.Pop();

            // is a var/fct call name?
            if (!InstrUtils.CheckObjectName(result, listVar, instrBase, out InstrNameObject instrNameObject))
                return false;

            if(instrNameObject!=null)
                instrBase = instrNameObject;

            listInstrBaseOut.Add(instrBase);
            return true;
        }

        // save instr in a list of instr in the right order
        List<InstrBase> listInstr= stackInstr.RemoveSaveInListReverse(instrCount);

        List<InstrBase> listOperand = new List<InstrBase>();
        List<InstrBase> listOperator = new List<InstrBase>();
        int i = 0;

        int commaCount = 0;
        // scan instr one by one
        while (i< listInstr.Count) 
        {
            if(!ProcessInstr(listInstr[i], listOperand, listOperator, out int commaCountLocal))
            {
                result.AddError(ErrorCode.ParserExpressionWrong, scriptToken);
                return false;
            }
            commaCount += commaCountLocal;
            i++;
        }
        // one or more error has occured
        if (!result.Res)
            return false;

        // no operand found, error
        if (listOperand.Count == 0)
        {
            result.AddError(ErrorCode.ParserExpressionWrong, listInstr[0].FirstScriptToken());
            return false;
        }

        // no operator and no comma found, error
        if (listOperator.Count == 0 && commaCount==0)
        {
            result.AddError(ErrorCode.ParserExpressionWrong, listInstr[0].FirstScriptToken());
            return false;
        }

        //--is it a comparison expression ?
        if (!IsInstrExprComparison(result, listOperator, listOperand, commaCount, out InstrComparison instrComparison))
            return false;
        if (instrComparison != null)
        {
            listInstrBaseOut.Add(instrComparison);
            return true;
        }

        //--is it a boolean expression ?
        if (!IsInstrBoolExpr(result, listOperator, listOperand, commaCount, out InstrBoolExpr instrBoolExpr))
            return false;
        if (instrBoolExpr != null)
        {
            listInstrBaseOut.Add(instrBoolExpr);
            return true;
        }

        //--is it a calculation expression ?
        // TODO:

        //--is it a function call list of parameters ?
        if (commaCount>0 && listOperator.Count==0 && listOperand.Count>0)
        {
            // TODO: need to check the sequence: operand comma,...
            if(listOperand.Count != commaCount+1)
            {
                result.AddError(ErrorCode.ParserTokenExpected, scriptToken);
                return false;
            }

            listInstrBaseOut = listOperand;
            return true;
        }

        // error, unable to parse the expression
        result.AddError(ErrorCode.ParserTokenExpected, scriptToken);
        return false;
    }

    /// <summary>
    /// Process the instruction.
    /// Is it a comparison operator ? a calculation operator? a comma? ...
    /// </summary>
    /// <param name="instr"></param>
    /// <param name="listOperand"></param>
    /// <param name="listOperator"></param>
    /// <param name="commaCount"></param>
    /// <returns></returns>
    static bool ProcessInstr(InstrBase instr, List<InstrBase> listOperand, List<InstrBase> listOperator, out int commaCount)
    {
        commaCount = 0;

        // the instr is an comparison operator?  =, >, <,...
        InstrSepComparison instrSepComparison = instr as InstrSepComparison;
        if (instrSepComparison != null)
        {
            listOperator.Add(instr);
            return true;
        }

        // the instr is a boolean operator?  and, or
        if (instr is InstrAnd || instr is InstrOr) 
        { 
            listOperator.Add(instr);
            return true;
        }

        // the instr is a calculation operator?  +, -, *, /
        if (instr is InstrCharMinus || instr is InstrCharPlus || instr is InstrCharMul || instr is InstrCharDiv)
        {
            listOperator.Add(instr);
            return true;
        }

        if(instr is InstrComma)
        {
            commaCount++;
            return true;
        }

        // a value?
        InstrValue instrValue = instr as InstrValue;
        if(instrValue!=null)
        {
            listOperand.Add(instr);
            return true;
        }

        // an object name?
        InstrNameObject instrNameObject = instr as InstrNameObject;
        if(instrNameObject!=null)
        {
            listOperand.Add(instr);
            return true;
        }

        // A.Cell?
        InstrColCellFunc instrColCellFunc = instr as InstrColCellFunc;
        if(instrColCellFunc!=null)
        {
            listOperand.Add(instr);
            return true;
        }

        // instr blank? instr null?
        if(instr.InstrType== InstrType.InstrBlank || instr.InstrType== InstrType.InstrNull)
        {
            listOperand.Add(instr);
            return true;
        }

        // a fct call?
        if (instr.IsFunctionCall)
        {
            listOperand.Add(instr);
            return true;
        }

        // instr comparison?
        InstrComparison instrComparison = instr as InstrComparison;
        if(instrComparison!=null)
        {
            listOperand.Add(instr);
            return true;
        }

        // instr bool expr ?
        InstrBoolExpr instrBoolExpr = instr as InstrBoolExpr;
        if (instrBoolExpr != null)
        {
            listOperand.Add(instr);
            return true;
        }

        return false;
    }


    /// <summary>
    ///  is it a comparison expression ?
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listOperator"></param>
    /// <param name="listOperand"></param>
    /// <param name="commaCount"></param>
    /// <param name="instrComparison"></param>
    /// <returns></returns>
    static bool IsInstrExprComparison(Result result, List<InstrBase> listOperator, List<InstrBase> listOperand, int commaCount, out InstrComparison instrComparison)
    {
        instrComparison = null;

        if (listOperator.Count == 0) return true;
        
        InstrSepComparison instrSepComparison = listOperator[0] as InstrSepComparison;
        if (instrSepComparison == null) return true;

        if (listOperator.Count > 1 || listOperand.Count > 2 || commaCount > 0)
        {
            result.AddError(ErrorCode.ParserCompExprWrong, instrSepComparison.FirstScriptToken());
            return false;
        }

        // check special case: blank and null, only equal or diff operator are allowed
        if(listOperand[0] is InstrBlank || listOperand[1] is InstrBlank || listOperand[0] is InstrNull || listOperand[1] is InstrNull)
        {
            if(instrSepComparison.Operator != SepComparisonOperator.Equal && instrSepComparison.Operator != SepComparisonOperator.Different)
            {
                result.AddError(ErrorCode.ParserSepComparatorWrong, instrSepComparison.FirstScriptToken());
                return false;
            }
        }

        // check special case: string, only equal or diff operator are allowed
        if (InstrUtils.IsValueString(listOperand[0]) || InstrUtils.IsValueString(listOperand[1]))
        {
            if (instrSepComparison.Operator != SepComparisonOperator.Equal && instrSepComparison.Operator != SepComparisonOperator.Different)
            {
                result.AddError(ErrorCode.ParserSepComparatorWrong, instrSepComparison.FirstScriptToken());
                return false;
            }
        }

        instrComparison = new InstrComparison(listOperand[0].FirstScriptToken());
        instrComparison.Operator = listOperator[0] as InstrSepComparison;
        instrComparison.OperandLeft = listOperand[0];
        instrComparison.OperandRight = listOperand[1];
        return true;
    }

    /// <summary>
    /// is it a bool expression ?
    /// exp: A.Cell>10 And A.Cell<20
    /// </summary>
    /// <param name="result"></param>
    /// <param name="listOperator"></param>
    /// <param name="listOperand"></param>
    /// <param name="commaCount"></param>
    /// <param name="instrBoolExpr"></param>
    /// <returns></returns>
    static bool IsInstrBoolExpr(Result result, List<InstrBase> listOperator, List<InstrBase> listOperand, int commaCount, out InstrBoolExpr instrBoolExpr)
    {
        instrBoolExpr = null;

        if (listOperator.Count == 0) return true;
        // 2 operand minimum
        if (listOperand.Count < 2) return true;

        bool isAnd = false;
        bool isOr = false;

        // should be And or Or
        foreach (InstrBase instrBoolOper in listOperator)
        {
            if(instrBoolOper is InstrAnd)
            {
                if (isOr)
                {
                    // pb, only one operator can be find, ANd or Or
                    result.AddError(ErrorCode.ParserBoolExprWrong, instrBoolOper.FirstScriptToken());
                    return false;
                }
                else  isAnd = true;
                continue;
            }

            if (instrBoolOper is InstrOr)
            {
                if (isAnd)
                {
                    // pb, only one operator can be find, ANd or Or
                    result.AddError(ErrorCode.ParserBoolExprWrong, instrBoolOper.FirstScriptToken());
                    return false;
                }
                else isOr = true;
                continue;
            }

            // operator not expected
            result.AddError(ErrorCode.ParserBoolExprWrong, instrBoolOper.FirstScriptToken());
            return false;
        }

        InstrBoolExprOperator instrBoolExprOperator = InstrBoolExprOperator.And;
        if (isOr) instrBoolExprOperator = InstrBoolExprOperator.Or;

        instrBoolExpr = new InstrBoolExpr(listOperand[0].FirstScriptToken(), instrBoolExprOperator);
        instrBoolExpr.ListOperand.AddRange(listOperand);    
        return true;
    }
}
