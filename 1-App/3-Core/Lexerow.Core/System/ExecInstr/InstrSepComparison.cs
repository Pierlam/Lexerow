using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System;

public enum SepComparisonOperator
{
    Undefined,
    Equal,
    Different,
    GreaterThan,
    LessThan,
    GreaterEqualThan,
    LessEqualThan,
}

public class InstrSepComparison : InstrBase
{
    public InstrSepComparison(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.SepComparison;

        if (scriptToken.Value == "=") Operator = SepComparisonOperator.Equal;
        if (scriptToken.Value == "<>") Operator = SepComparisonOperator.Different;
        if (scriptToken.Value == ">") Operator = SepComparisonOperator.GreaterThan;
        if (scriptToken.Value == "<") Operator = SepComparisonOperator.LessThan;
        if (scriptToken.Value == ">=") Operator = SepComparisonOperator.GreaterEqualThan;
        if (scriptToken.Value == "<=") Operator = SepComparisonOperator.LessEqualThan;
    }

    public SepComparisonOperator Operator { get; set; } = SepComparisonOperator.Undefined;

    public InstrSepComparison Revert()
    {
        InstrSepComparison instrSepComparison = new InstrSepComparison(this.FirstScriptToken());

        // < becomes >
        if (Operator == SepComparisonOperator.LessThan)
        {
            instrSepComparison.Operator = SepComparisonOperator.GreaterThan;
            return instrSepComparison;
        }

        // > becomes <
        if (Operator == SepComparisonOperator.GreaterThan)
        {
            instrSepComparison.Operator = SepComparisonOperator.LessThan;
            return instrSepComparison;
        }

        // = becomes =<
        if (Operator == SepComparisonOperator.GreaterEqualThan)
        {
            instrSepComparison.Operator = SepComparisonOperator.LessEqualThan;
            return instrSepComparison;
        }

        // =< becomes >=
        if (Operator == SepComparisonOperator.LessEqualThan)
        {
            instrSepComparison.Operator = SepComparisonOperator.GreaterEqualThan;
            return instrSepComparison;
        }
        // equals and different doesnt change
        return instrSepComparison;
    }
}