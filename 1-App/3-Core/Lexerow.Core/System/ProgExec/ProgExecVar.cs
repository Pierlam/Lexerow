using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.ProgExec;

public enum ProgRunVarType
{
    // string, int, double, ...
    BasicValue,

    ExcelFile

}

public class ProgExecVar
{
    public ProgExecVar(InstrBase name, InstrBase value)
    {
        ObjectName = name;
        Value = value;
    }

    /// <summary>
    /// The Name of the variable.
    /// The left part, can be an object.
    /// a=12        -> a      -> InstrObjectName
    /// A.Cell= 12  -> A.Cell -> InstrColCellFunc
    /// </summary>
    public InstrBase ObjectName { get; set; }

    /// <summary>
    /// The value of the variable.
    /// </summary>
    public InstrBase Value { get; set; }

    public bool NameEquals(string name)
    {
        InstrObjectName instrObjectName= ObjectName as InstrObjectName;
        if (instrObjectName != null) 
        {
            if (instrObjectName.ObjectName.Equals(name, StringComparison.InvariantCultureIgnoreCase)) return true;
        }
        return false;
    }

    public string GetValueString()
    {
        InstrConstValue instrConstValue = Value as InstrConstValue;
        if(instrConstValue!=null)
            return instrConstValue.RawValue;

        return string.Empty;
    }

    public bool AreSame(InstrBase instr)
    {
        if(ObjectName.InstrType!= instr.InstrType) return false;

        //--possible??
        if(ObjectName.InstrType== InstrType.ConstValue)
        {
            string name = (ObjectName as InstrConstValue).RawValue;
            string name2= (instr as InstrConstValue).RawValue;
            if (name.Equals(name2))return true;

            return false;            
        }

        //--manage var name, exp: file
        if (ObjectName.InstrType == InstrType.ObjectName)
        {
            string name = (ObjectName as InstrObjectName).ObjectName;
            string name2 = (instr as InstrObjectName).ObjectName;
            if (name.Equals(name2)) return true;

            return false;
        }

        //--Manage A.Cell format
        if (ObjectName.InstrType == InstrType.ColCellFunc)
        {
            // compare col name first
            // TODO:

            // then compare function
            // TODO:
            throw new Exception("Not yet implemented");
        }

        return false;
    }
}
