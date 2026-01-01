using Lexerow.Core.System;
using Lexerow.Core.System.ProgExec;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Utils;


/// <summary>
///  Some useful functions around ValueBase classes.
/// </summary>
public class ValueUtils
{
    /// <summary>
    /// Return if both value type match.
    /// </summary>
    /// <param name="value1"></param>
    /// <param name="value2"></param>
    /// <returns></returns>
    public static  bool TypeMatch(ValueBase value1, ValueBase value2)
    {
        if(value1==null || value2==null)  return false;

        // check that the new value match the type
        if(value1.ValueType == value2.ValueType) return true;
        return false;
    }

    /// <summary>
    /// Copy the value from (source) to the target value.
    /// ValueTo= ValueFrom
    /// !Check before that both type match!
    /// </summary>
    /// <param name="valueTo"></param>
    /// <param name="valueFrom"></param>
    /// <returns></returns>
    public static bool CopyValue(ValueBase valueTo, ValueBase valueFrom)
    {
        if (valueTo.ValueType == System.ValueType.String)
        {
            (valueTo as ValueString).Val = (valueFrom as ValueString).Val;
            return true; 
        }

        if (valueTo.ValueType == System.ValueType.Int)
        {
            (valueTo as ValueInt).Val = (valueFrom as ValueInt).Val;
            return true;
        }

        if (valueTo.ValueType == System.ValueType.Double)
        {
            (valueTo as ValueDouble).Val = (valueFrom as ValueDouble).Val;
            return true;
        }

        if (valueTo.ValueType == System.ValueType.DateTime)
        {
            (valueTo as ValueDateTime).Val = (valueFrom as ValueDateTime).Val;
            return true;
        }

        if (valueTo.ValueType == System.ValueType.DateOnly)
        {
            (valueTo as ValueDateOnly).Val = (valueFrom as ValueDateOnly).Val;
            return true;
        }

        if (valueTo.ValueType == System.ValueType.TimeOnly)
        {
            (valueTo as ValueTimeOnly).Val = (valueFrom as ValueTimeOnly).Val;
            return true;
        }

        if (valueTo.ValueType == System.ValueType.Bool)
        {
            (valueTo as ValueBool).Val = (valueFrom as ValueBool).Val;
            return true;
        }

        if (valueTo.ValueType == System.ValueType.ListOfString)
        {
            (valueTo as ValueListOfString).ListVal.Clear();
            (valueTo as ValueListOfString).ListVal.AddRange((valueFrom as ValueListOfString).ListVal);
            return true;
        }

        if (valueTo.ValueType == System.ValueType.ListOfDouble)
        {
            (valueTo as ValueListOfDouble).ListVal.Clear();
            (valueTo as ValueListOfDouble).ListVal.AddRange((valueFrom as ValueListOfDouble).ListVal);
            return true;
        }

        if (valueTo.ValueType == System.ValueType.ListOfInt)
        {
            (valueTo as ValueListOfInt).ListVal.Clear();
            (valueTo as ValueListOfInt).ListVal.AddRange((valueFrom as ValueListOfInt).ListVal);
            return true;
        }

        // not managed
        return false;
    }
}
