using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.InstrDef;

public enum InstrBoolExprOperator
{
    And,
    Or
}

/// <summary>
/// Instruction boolean expression.
/// Only one operator type: And or Or.
/// Can have many operand, each operand must return a bool value.
/// An operand can be a bool value, an comparison instr or a function call (returning a bool value).
/// exp: If A.Cell>10 And A.Cell<20 
/// exp: If MyFct()
/// exp: If a And A.Cell>10
/// </summary>
public class InstrBoolExpr : InstrBase
{
    public InstrBoolExpr(ScriptToken scriptToken, InstrBoolExprOperator oper) : base(scriptToken)
    {
        InstrType= InstrType.BoolExpr;
        Operator = oper;
    }

    /// <summary>
    /// The operator to apply between each operand: And or Or
    /// </summary>
    public InstrBoolExprOperator Operator { get; set; }

    public List<InstrBase> ListOperand { get; set; }=new List<InstrBase>();

    public override string ToString()
    {
        string s=string.Empty;
        foreach (var instr in ListOperand) 
        {
            s+=instr.ToString()+", ";
        }
        // remove the last comma
        if(s.Trim().EndsWith(",")) s=s.Trim().Substring(0, s.Length-1);
        return "BoolExpr, oper: " + Operator.ToString()+", operands: " + s;
    }
}
