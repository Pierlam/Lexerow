using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Compilator;
/// <summary>
/// obtained by the paser, first stage: lexical analyse.
/// Split the code source line by line to obtains tokens.
/// </summary>
public enum SourceCodeTokenType
{
    Undefined,

    /// <summary>
    /// An item can be: a variable, object, function or method name.
    /// </summary>
    Name,


    /// <summary>
    /// excel column name.
    /// exp: A, VB,XFD
    /// </summary>
    ExcelColName,

    /// <summary>
    /// Excel cell address.
    /// exp: A1 , BN23
    /// </summary>
    ExcelCellAddress,

    /// <summary>
    /// Separator
    /// exp: =,;.< >()+-*/  ...
    /// </summary>
    Separator,

    /// <summary>
    /// Wellformed string,has start and end tags.
    /// </summary>
    String,

    /// <summary>
    /// has the first string tag, but not the last.
    /// </summary>
    StringBadFormed,

    Integer,

    Double,

    /// <summary>
    /// has for example two or more decimal separator.
    /// </summary>
    DoubleWrong
}

public class SourceCodeToken
{
    /// <summary>
    /// Line Number in the source code.
    /// </summary>
    public int LineNum { get; set; } = 0;

    /// <summary>
    /// Column number in the source code line.
    /// Index first char of the token.
    /// base0
    /// </summary>
    public int ColNum { get; set; }

    /// <summary>
    /// the token value.
    /// </summary>
    public string Value { get; set; }

    public SourceCodeTokenType SourceCodeTokenType { get; set; }

    /// <summary>
    /// Can be an integer, double, or a string.
    /// </summary>
    /// <returns></returns>
    public bool IsConstValue()
    {
        // is the token a const value?
        if (SourceCodeTokenType == SourceCodeTokenType.Integer || SourceCodeTokenType == SourceCodeTokenType.String ||
            SourceCodeTokenType == SourceCodeTokenType.StringBadFormed || SourceCodeTokenType == SourceCodeTokenType.Double ||
            SourceCodeTokenType == SourceCodeTokenType.DoubleWrong)
            return true;

        return false;
    }
}
