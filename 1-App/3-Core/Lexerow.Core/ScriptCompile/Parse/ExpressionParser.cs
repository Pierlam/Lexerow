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
/// Parse expression found between mainly If and Then, or bertween ( and ).
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
    public static bool Process(Result result, List<InstrNameObject> listVar, CompilStackInstr stackInstr, ScriptToken scriptToken, InstrType instrTypeStart, out InstrBase instrBaseOut)
    {
        instrBaseOut = null;

        // get instr until the start instr: If or (
        int instrCount = stackInstr.GetDistanceFromTop(instrTypeStart) - 1;

        //// even or odd?
        //bool isEven = instrCount % 2 == 0;
        //// wrong count of instr: 0, 2, 4, 6,...
        //if (instrCount == 0 || isEven)
        //{
        //    result.AddError(ErrorCode.ParserTokenExpected, scriptToken);
        //    return false;
        //}

        //--before, on stack, is it a bool value? exp: If a Then, If (a) Then or a fct returning a bool value, fct()
        if (instrCount ==1)
        {
            InstrBase instrBase = stackInstr.Pop();
            if (instrBase.ReturnType== InstrReturnType.ValueBool)
            {
                instrBaseOut = instrBase;
                return true;
            }
        }

        // save instr in a list of instr in the right order
        List<InstrBase> listInstr= stackInstr.RemoveSaveInListReverse(instrCount);
        bool isCompExpr = false;

        List<InstrBase> listOperand = new List<InstrBase>();
        List<InstrBase> listOperator = new List<InstrBase>();
        int i = 0;

        // scan instr one by one
        while (i< listInstr.Count) 
        {
            ProcessInstr(listInstr[i], listOperand, listOperator, out bool isCompExprLocal);
            isCompExpr= isCompExprLocal;
            i++;
        }

        // no operand found, error
        if (listOperand.Count == 0)
        {
            result.AddError(ErrorCode.ParserExpressionWrong, listInstr[0].FirstScriptToken());
            return false;
        }

        // no operator found, error
        if (listOperator.Count == 0)
        {
            result.AddError(ErrorCode.ParserExpressionWrong, listInstr[0].FirstScriptToken());
            return false;
        }

        //--is it a comparison expression ?
        InstrSepComparison instrSepComparison = listOperator[0] as InstrSepComparison;
        if (instrSepComparison != null) 
        {
            if (listOperator.Count > 1 || listOperand.Count > 2)
            {
                result.AddError(ErrorCode.ParserCompExprWrong, listInstr[0].FirstScriptToken());
                return false;
            }
            InstrComparison instrComparison=new InstrComparison(listOperand[0].FirstScriptToken());
            instrComparison.Operator = listOperator[0] as InstrSepComparison;
            instrComparison.OperandLeft = listOperand[0];
            instrComparison.OperandRight = listOperand[1];

            instrBaseOut = instrComparison;
            return true;
        }

        //--is it a boolean expression ?
        // TODO:

        //--is it a calculation expression ?
        // TODO:

        //--is it a function call list of parameters ?
        // TODO:

        // error, unable to parse the expression
        result.AddError(ErrorCode.ParserTokenExpected, scriptToken);
        return false;
    }

    static bool ProcessInstr(InstrBase instr, List<InstrBase> listOperand, List<InstrBase> listOperator, out bool isCompExpr)
    {
        isCompExpr= false;

        // the instr is an comparison operator?  =, >, <,...
        InstrSepComparison instrSepComparison = instr as InstrSepComparison;
        if (instrSepComparison != null)
        {
            listOperator.Add(instr);
            isCompExpr = true;
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

        // the instr is an operand?  
        // TODO: check it!
        listOperand.Add(instr);

        return true;
    }
}
