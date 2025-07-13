using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Lexerow.Core.Scripts.Compilator;
public class StringParser
{
    /// <summary>
    /// Parse a string source code, split it in items: separator, string, token.
    /// remove each comment.
    /// </summary>
    /// <param name="line"></param>
    /// <param name="separators"></param>
    /// <param name="stringSep"></param>
    /// <param name="commentTag"></param>
    /// <param name="tokens"></param>
    /// <returns></returns>
    public bool Parse(int lineNum, string line, string separators, char stringSep, string commentTag, out List<SourceCodeToken> tokens)
    {
        tokens = new List<SourceCodeToken>();
        SourceCodeToken token;
        int i = 0;
        int iOut;

        if (string.IsNullOrEmpty(line))
            return true;

        while (true)
        {
            // no more char to extract
            if (i >= line.Length)
            {
                break;
            }

            //--is there some space separator?
            if (ProcessSpaceChars(line, i, out iOut))
            {
                i = iOut;
                continue;
            }

            //--is the token a 'string' ?
            if (ProcessString(line, i, stringSep, out iOut, out token))
            {
                tokens.Add(token);
                i = iOut;
                continue;
            }

            //--is it a comment ?
            if (ProcessComment(line, i, commentTag))
            {
                // the comment atg is found, the rest ofthe string is a comment
                return true;
            }

            //--is is a number?, integer or double, should be before separators extraction! because of decimal separator
            if (ProcessNumber(line, i, out iOut, out token))
            {
                tokens.Add(token);
                i = iOut;
                continue;
            }

            //--is is a char separator?
            if (ProcessSeparator(line, i, separators, out iOut, out token))
            {
                tokens.Add(token);
                i = iOut;
                continue;
            }

            //--is it a name? variable, constant, object, function  or method
            if (ProcessName(line, i, out iOut, out token))
            {
                // is it an excel column name?
                ProcessExcelColName(token);

                //--is an Excel cell address?
                ProcessExcelCellAddress(token);

                tokens.Add(token);
                i = iOut;
                continue;
            }

            // if here, unexpected char!
            token = new SourceCodeToken();
            token.SourceCodeTokenType = SourceCodeTokenType.Undefined;
            token.LineNum = lineNum;
            token.ColNum = i;
            token.Value = line[i].ToString();
            tokens.Add(token);
            i++;
        }

        return true;
    }

    /// <summary>
    /// Move cursor on each space char found.
    /// </summary>
    /// <param name="line"></param>
    /// <param name="i"></param>
    /// <param name="iOut"></param>
    bool ProcessSpaceChars(string line, int i, out int iOut)
    {
        bool spaceFound = false;

        iOut = i;
        while (true)
        {
            if (i >= line.Length) return spaceFound;
            if (line[i] != ' ') return spaceFound;

            spaceFound = true;
            i++;
            iOut = i;
        }
    }

    bool ProcessString(string line, int i, char stringSep, out int iOut, out SourceCodeToken item)
    {
        iOut = i;
        item = null;

        // is it a string separator?
        char c = line[i];
        if (c != stringSep)
            return false;

        // it's a string
        item = new SourceCodeToken();
        item.SourceCodeTokenType = SourceCodeTokenType.String;
        item.Value = stringSep.ToString();
        item.ColNum = i;

        // find the end string separator
        while (true)
        {
            i++;

            if (i >= line.Length)
            {
                iOut = i;
                item.SourceCodeTokenType = SourceCodeTokenType.StringBadFormed;
                return true;
            }

            c = line[i];
            item.Value += c.ToString();

            // is it a string separator?
            if (c == stringSep)
            {
                // move to the next char after the string sep
                i++;
                iOut = i;
                return true;
            }
        }
    }

    bool ProcessComment(string line, int i, string commentTag)
    {
        if (line.Substring(i).StartsWith(commentTag, StringComparison.OrdinalIgnoreCase))
            return true;

        return false;
    }

    /// <summary>
    /// find numbers: integer and double. exp: 12  12.34
    /// Decimal separator is: .
    /// </summary>
    /// <param name="line"></param>
    /// <param name="i"></param>
    /// <param name="iOut"></param>
    /// <param name="token"></param>
    /// <returns></returns>
    bool ProcessNumber(string line, int i, out int iOut, out SourceCodeToken token)
    {
        iOut = i;
        token = null;
        bool tokenFound = false;

        while (true)
        {
            if (i >= line.Length) return tokenFound;

            char c = line[i];

            // is it a letter or a digit or an underscore?
            if (char.IsDigit(c))
            {
                if (token == null)
                {
                    token = new SourceCodeToken();
                    token.SourceCodeTokenType = SourceCodeTokenType.Integer;
                    token.ColNum = i;
                }

                token.Value += c.ToString();
                i++;
                iOut = i;
                tokenFound = true;
                continue;
            }

            // is it the decimal separator?
            if (c == '.' && tokenFound)
            {
                // contains already a decimal separator!
                if (token.Value.Contains("."))
                    token.SourceCodeTokenType = SourceCodeTokenType.DoubleWrong;
                else
                    token.SourceCodeTokenType = SourceCodeTokenType.Double;

                token.Value += c.ToString();
                i++;
                iOut = i;
                continue;
            }

            // not a digit or a decimal separator!
            //ici();

            return tokenFound;
        }
    }

    bool ProcessSeparator(string line, int i, string separators, out int iOut, out SourceCodeToken token)
    {
        iOut = i;
        token = null;

        if (separators.Contains(line[i]))
        {
            // it's a separator
            token = new SourceCodeToken();
            token.SourceCodeTokenType = SourceCodeTokenType.Separator;
            token.Value = line[i].ToString();
            token.ColNum = i;
            iOut = i + 1;
            return true;
        }
        return false;
    }

    /// <summary>
    /// Is it a name?
    /// variable, object, function, or method.
    /// </summary>
    /// <param name="line"></param>
    /// <param name="i"></param>
    /// <param name="iOut"></param>
    /// <param name="token"></param>
    /// <returns></returns>
    bool ProcessName(string line, int i, out int iOut, out SourceCodeToken token)
    {
        iOut = i;
        token = null;
        bool tokenFound = false;
        while (true)
        {
            if (i >= line.Length) return tokenFound;

            char c = line[i];

            // is it a letter or a digit or an underscore?
            if (char.IsLetterOrDigit(c) || c == '_')
            {
                if (token == null)
                {
                    token = new SourceCodeToken();
                    token.SourceCodeTokenType = SourceCodeTokenType.Name;
                    token.ColNum = i;
                }

                token.Value += c.ToString();
                i++;
                iOut = i;
                tokenFound = true;
                continue;
            }

            return tokenFound;
        }
    }

    /// <summary>
    /// Is the token an excel column name?
    /// exp: A
    /// max: XFD
    /// </summary>
    /// <param name="line"></param>
    /// <param name="i"></param>
    /// <param name="separators"></param>
    /// <param name="iOut"></param>
    /// <param name="token"></param>
    /// <returns></returns>
    bool ProcessExcelColName(SourceCodeToken token)
    {
        int i = 0;
        bool found = false;

        // find the end string separator
        while (true)
        {
            // stop until end of line reached or a separator char is found
            if (i >= token.Value.Length) break;

            // the char should be:letter+upper
            if (char.IsLetter(token.Value[i]) && char.IsUpper(token.Value[i]))
                found = true;
            else
            {
                found = false;
                break;
            }

            i++;
        }

        if (found)
        {
            token.SourceCodeTokenType = SourceCodeTokenType.ExcelColName;
        }
        return found;
    }

    /// <summary>
    /// Is the token an excel cell address?
    /// exp: A1, BN456
    /// max: XFD???
    /// 
    /// </summary>
    /// <param name="token"></param>
    /// <returns></returns>
    bool ProcessExcelCellAddress(SourceCodeToken token)
    {

        string pattern = @"(\$?[A-Z]+\$?\d+)";
        Regex regex = new Regex(pattern);
        //MatchCollection matches = regex.Matches(token.Value);
        if (regex.Match(token.Value).Success)
        {
            token.SourceCodeTokenType = SourceCodeTokenType.ExcelCellAddress;
            return true;
        }

        return false;
    }
}
