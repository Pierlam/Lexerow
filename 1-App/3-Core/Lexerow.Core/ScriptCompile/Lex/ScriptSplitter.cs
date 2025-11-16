using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Lexerow.Core.ScriptCompile.lex;
public class ScriptSplitter
{
    /// <summary>
    /// decimal separator. exp: 10.34
    /// </summary>
    public char DecimalSep = '.';

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
    public bool Split(int lineNum, string line, string separators, char stringSep, string commentTag, out List<ScriptToken> tokens, out ScriptTokenType lastTokenType)
    {
        lastTokenType = ScriptTokenType.Comment;
        tokens = new List<ScriptToken>();
        ScriptToken token;
        int i = 0;
        int iOut;

        if (string.IsNullOrEmpty(line))
        {
            return true;
        }

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
                lastTokenType = ScriptTokenType.Comment;
                i = iOut;
                continue;
            }

            //--is the token a 'string' ?
            if (ProcessString(line, i, stringSep, out iOut, out token))
            {
                lastTokenType = ScriptTokenType.String;
                token.LineNum = lineNum;
                tokens.Add(token);
                i = iOut;

                // manage case: StringBadFormated, end tag not found
                if (token.ScriptTokenType == ScriptTokenType.StringBadFormed) 
                {
                    lastTokenType = ScriptTokenType.StringBadFormed;
                    return false;
                }
                continue;
            }

            //--is it a comment ?
            if (ProcessComment(line, i, commentTag))
            {
                lastTokenType = ScriptTokenType.Comment;
                // the comment tag is found, the rest ofthe string is a comment
                return true;
            }

            //--is is a number?, integer or double, should be before separators extraction! because of decimal separator
            if (ProcessNumber(separators, DecimalSep, line, i, out iOut, out token))
            {
                token.LineNum = lineNum;
                tokens.Add(token);
                i = iOut;

                // manage case: StringBadFormated, end tag not found
                if (token.ScriptTokenType == ScriptTokenType.WrongNumber)
                {
                    lastTokenType = ScriptTokenType.WrongNumber;
                    return false;
                }
                continue;
            }

            //--is is a char separator?  manage special cases: >=, <=, <>
            if (ProcessSeparator(separators, line, i, out iOut, out token))
            {
                token.LineNum = lineNum;
                tokens.Add(token);
                i = iOut;
                lastTokenType = ScriptTokenType.Separator;

                continue;
            }

            //--is it a name? variable, constant, object, function  or method
            if (ProcessName(line, i, out iOut, out token))
            {
                token.LineNum = lineNum;
                tokens.Add(token);
                lastTokenType = ScriptTokenType.Name;
                i = iOut;
                continue;
            }

            // if here, unexpected char!
            token = new ScriptToken();
            token.ScriptTokenType = ScriptTokenType.Undefined;
            token.LineNum = lineNum;
            token.ColNum = i;
            token.Value = line[i].ToString();
            tokens.Add(token);
            i++;

            lastTokenType = ScriptTokenType.Undefined;
            return false;
        }

        return true;
    }

    /// <summary>
    /// Move cursor on each space char found.
    /// can be: space, \n, \t, \r.
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

            // space or special char (\n,\r, \t,...) found?
            if(!IsSpaceCharExt(line[i]))return spaceFound;

            spaceFound = true;
            i++;
            iOut = i;
        }
    }

    bool ProcessString(string line, int i, char stringSep, out int iOut, out ScriptToken item)
    {
        iOut = i;
        item = null;

        // is it a string separator?
        char c = line[i];
        if (c != stringSep)
            return false;

        // it's a string
        item = new ScriptToken();
        item.ScriptTokenType = ScriptTokenType.String;
        item.Value = stringSep.ToString();
        item.ColNum = i;

        // find the end string separator
        while (true)
        {
            i++;

            if (i >= line.Length)
            {
                iOut = i;
                item.ScriptTokenType = ScriptTokenType.StringBadFormed;
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
    /// Decimal separator is the dot: .
    /// 12E1O, 23E-10
    /// </summary>
    /// <param name="line"></param>
    /// <param name="i"></param>
    /// <param name="iOut"></param>
    /// <param name="token"></param>
    /// <returns></returns>
    bool ProcessNumber(string separators, char decimalSep, string line, int i, out int iOut, out ScriptToken token)
    {
        iOut = i;
        token = null;
        bool tokenFound = false;

        while (true)
        {
            // no more char
            if (i >= line.Length) break;

            char c = line[i];

            if (char.IsDigit(c))
            {
                if (token == null)
                {
                    token = new ScriptToken();
                    token.ScriptTokenType = ScriptTokenType.Integer;
                    token.ColNum = i;
                }

                token.Value += c.ToString();
                i++;
                iOut = i;
                tokenFound = true;
                continue;
            }

            // is it the decimal separator?
            if (c == decimalSep)
            {
                if (token == null)
                {
                    // the next char should exists and should be a digit
                    if (i+1 < line.Length) 
                    {
                        // next char must be a digit
                        if (char.IsDigit(line[i + 1]))
                        {
                            token = new ScriptToken();
                            token.ScriptTokenType = ScriptTokenType.Double;
                            token.ColNum = i;
                            continue;
                        }
                    }

                    // not a double number!
                    break;
                }

                // sure it's a double
                token.ScriptTokenType = ScriptTokenType.Double;

                token.Value += c.ToString();
                i++;
                iOut = i;
                continue;
            }

            // is it the E exposant separator ?
            if (c == 'e' || c == 'E')
            {
                // the token should exists
                if (token != null) 
                {
                    // sure it's a double
                    token.ScriptTokenType = ScriptTokenType.Double;

                    token.Value += c.ToString();
                    i++;
                    iOut = i;
                    continue;
                }
            }

            // special case: minus char, exp 23E-10
            if(c=='-')
            {
                if(token != null)
                {
                    // previous char should exists and should be: E
                    if(token.Value.Length > 0)
                    {
                        if(token.Value.Last() == 'E' || token.Value.Last() == 'e')
                        {
                            token.Value += c.ToString();
                            i++;
                            iOut = i;
                            continue;
                        }
                    }
                }
            }

            // not a digit or a decimal separator!
            break;
        }

        // no integer or double found, bye
        if (token==null)
            return false;

        // the next char must be a separator or there is no more char
        if(iOut < line.Length)
        {
            // there is a next char, should be a space separator
            if (!IsSpaceCharExt(line[i]) && !separators.Contains(line[i]))
            {
                // error!  -> voir plus haut faire une fct
                // TODO: space, \n,\r
                token.ScriptTokenType = ScriptTokenType.WrongNumber;
                return true;
            }
        }

        // check the number, convert it
        if(token.ScriptTokenType == ScriptTokenType.Double)
        {
            if (double.TryParse(token.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
                token.ValueDouble = d;
            else
                token.ScriptTokenType= ScriptTokenType.WrongNumber;
            return true;
        }

        if (token.ScriptTokenType == ScriptTokenType.Integer)
        {
            if (int.TryParse(token.Value, out int d))
                token.ValueInt = d;
            else
                token.ScriptTokenType = ScriptTokenType.WrongNumber;
            return true;
        }

        // a number was found
        return true;
    }

    /// <summary>
    /// is the current char a separator?  .=,;+- 
    /// Manage special cases: >=, <=, <>
    /// </summary>
    /// <param name="separators"></param>
    /// <param name="line"></param>
    /// <param name="i"></param>
    /// <param name="iOut"></param>
    /// <param name="token"></param>
    /// <returns></returns>
    bool ProcessSeparator(string separators, string line, int i, out int iOut, out ScriptToken token)
    {
        iOut = i;
        token = null;

        if (separators.Contains(line[i]))
        {
            // it's a separator
            token = new ScriptToken();
            token.ScriptTokenType = ScriptTokenType.Separator;
            token.Value = line[i].ToString();
            token.ColNum = i;
            // next
            i++;
            iOut = i;

            // special case:  >=, <=, <>
            ProcessLessGreaterEqualSeparator(line, i, token, out iOut);
            return true;
        }
        return false;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="line"></param>
    /// <param name="i"></param>
    /// <param name="token"></param>
    /// <param name="iOut"></param>
    /// <returns></returns>
    bool ProcessLessGreaterEqualSeparator(string line, int i, ScriptToken token, out int iOut)
    {
        iOut = i;
        // no more char to process
        if (i > line.Length) return true;

        // >=
        if(token.Value==">" && line[i]=='=')
        {
            token.Value = ">=";
            iOut = i + 1;
            return true;
        }

        // <=
        if (token.Value == "<" && line[i] == '=')
        {
            token.Value = "<=";
            iOut = i + 1;
            return true;
        }

        // <>
        if (token.Value == "<" && line[i] == '>')
        {
            token.Value = "<>";
            iOut = i + 1;
            return true;
        }
        return true;
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
    bool ProcessName(string line, int i, out int iOut, out ScriptToken token)
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
                    token = new ScriptToken();
                    token.ScriptTokenType = ScriptTokenType.Name;
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
    /// Is the char a pure separator char like space?
    /// </summary>
    /// <param name="c"></param>
    /// <returns></returns>
    bool IsSpaceCharExt(char c)
    {
        if (c == ' ') return true;
        if (c == '\r') return true;
        if (c == '\n') return true;
        if (c == '\t') return true;
        return false;
    }
}
