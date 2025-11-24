namespace Lexerow.Core.System.GenDef;

public class CoreInstr
{    
    public const string InstrOnExcel = "OnExcel";
    public const string InstrOnSheet = "OnSheet";
    public const string InstrFirstRow = "FirstRow";
    public const string InstrForEach = "ForEach";

    public const string InstrRow = "Row";

    // special case
    public const string InstrForEachRow = "ForEachRow";

    public const string InstrNext = "Next";

    public const string InstrCol = "Col";
    public const string InstrCell = "Cell";

    public const string InstrIf = "If";
    public const string InstrThen = "Then";
    public const string InstrElse = "Else";

    public const string InstrBlank = "Blank";
    public const string InstrNull = "Null";
    public const string InstrEndName = "End";

    //--function names
    public const string InstrFuncSelectFiles = "SelectFiles";
    public const string InstrFuncDate = "Date";
    public const string InstrFuncDateTime = "DateTime";
    public const string InstrFuncTime = "Time";

    /// <summary>
    /// Used in OnExcel instr.
    /// The first data row to process.
    /// (after the header row)
    /// 
    /// value in base1 (human readable).
    /// </summary>
    public static int FirstDataRowNum = 2;
}