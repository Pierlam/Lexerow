using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Compilator;

public enum SyntaxAnalyserItemType
{
    // TODO: rework!!
    SourceToken,
    Instruction
}

// TODO: vraiment besoin??
public class SyntaxAnalyserItem
{

    public static SyntaxAnalyserItem CreateInstr(ExecTokBase instrBase)
    {
        SyntaxAnalyserItem item = new SyntaxAnalyserItem();
        item.ExecTokBase = instrBase;
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
    /// must be set.
    /// A script token is always converted to an exec token.
    /// </summary>
    public ExecTokBase? ExecTokBase { get; set; } = null;

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
        return (ExecTokBase as InstrSetVar) != null;
    }

    public bool IsInstrIfThenElse()
    {
        if (Type != SyntaxAnalyserItemType.Instruction) return false;
        return (ExecTokBase as InstrIfThenElse) != null;
    }

    public bool IsInstrCallMember()
    {
        return (ExecTokBase as InstrCallMemberBase) != null;
    }
}
