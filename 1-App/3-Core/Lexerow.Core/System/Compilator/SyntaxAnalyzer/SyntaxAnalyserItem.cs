using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Compilator;

public enum SyntaxAnalyserItemType
{
    SourceToken,
    Instruction
}

public class SyntaxAnalyserItem
{
    public static SyntaxAnalyserItem CreateToken(ScriptToken sourceCodeToken)
    {
        SyntaxAnalyserItem item = new SyntaxAnalyserItem();
        item.SourceCodeToken = sourceCodeToken;
        item.Type = SyntaxAnalyserItemType.SourceToken;

        return item;
    }

    public static SyntaxAnalyserItem CreateInstr(InstrBase instrBase)
    {
        SyntaxAnalyserItem item = new SyntaxAnalyserItem();
        item.InstrBase = instrBase;
        item.Type = SyntaxAnalyserItemType.Instruction;

        return item;
    }

    // type: Instr or token
    public SyntaxAnalyserItemType Type { get; set; } = SyntaxAnalyserItemType.SourceToken;

    /// <summary>
    /// The item can be a source code token.
    /// </summary>
    public ScriptToken? SourceCodeToken { get; set; } = null;

    /// <summary>
    /// The item can be a, instruction.
    /// </summary>
    public InstrBase? InstrBase { get; set; } = null;

    public bool IsTokenConstValue()
    {
        if (Type != SyntaxAnalyserItemType.SourceToken)
            // not a token, it's an instruction, bye
            return false;

        return SourceCodeToken.IsConstValue();
    }

    public bool IsTokenInteger()
    {
        if (Type != SyntaxAnalyserItemType.SourceToken)
            // not a token, it's an instruction, bye
            return false;

        return SourceCodeToken.ScriptTokenType == ScriptTokenType.Integer;

    }
    public bool IsTokenString()
    {
        if (Type != SyntaxAnalyserItemType.SourceToken)
            // not a token, it's an instruction, bye
            return false;

        return SourceCodeToken.ScriptTokenType == ScriptTokenType.String;

    }

    public bool IsTokenBracketClose()
    {
        if (Type != SyntaxAnalyserItemType.SourceToken)
            // not a token, it's an instruction, bye
            return false;

        return SourceCodeToken.Value == ")";
    }


    public bool IsTokenBracketOpen()
    {
        if (Type != SyntaxAnalyserItemType.SourceToken)
            // not a token, it's an instruction, bye
            return false;

        return SourceCodeToken.Value == "(";
    }

    /// <summary>
    /// fct parameters separator.
    /// </summary>
    /// <returns></returns>
    public bool IsTokenComma()
    {
        if (Type != SyntaxAnalyserItemType.SourceToken)
            // not a token, it's an instruction, bye
            return false;

        return SourceCodeToken.Value == ",";
    }

    public bool IsTokenVarName()
    {
        if (Type != SyntaxAnalyserItemType.SourceToken)
            // not a token, iti's an instruction, bye
            return false;

        return SourceCodeToken.ScriptTokenType == ScriptTokenType.Name;
    }

    //public bool IsTokenExcelCellAddress()
    //{
    //    if (Type != SyntaxAnalyserItemType.SourceToken)
    //        // not a token, iti's an instruction, bye
    //        return false;

    //    return SourceCodeToken.ScriptTokenType == ScriptTokenType.ExcelCellAddress;
    //}

    //public bool IsTokenExcelColName()
    //{
    //    if (Type != SyntaxAnalyserItemType.SourceToken)
    //        // not a token, iti's an instruction, bye
    //        return false;

    //    return SourceCodeToken.ScriptTokenType == ScriptTokenType.ExcelColName;
    //}

    /// <summary>
    /// is it an instruction?
    /// std or callMember.
    /// </summary>
    /// <returns></returns>
    public bool IsInstr()
    {
        return Type == SyntaxAnalyserItemType.Instruction;
    }

    public bool IsInstrSetVar()
    {
        if (Type != SyntaxAnalyserItemType.Instruction) return false;
        return (InstrBase as InstrSetVar) != null;
    }

    public bool IsInstrIfThenElse()
    {
        if (Type != SyntaxAnalyserItemType.Instruction) return false;
        return (InstrBase as InstrIfThenElse) != null;
    }

    public bool IsInstrCallMember()
    {
        return (InstrBase as InstrCallMemberBase) != null;
    }
}
