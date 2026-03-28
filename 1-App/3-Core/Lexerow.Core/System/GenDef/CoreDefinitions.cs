namespace Lexerow.Core.System.GenDef;

/// <summary>
/// Main core definitions.
/// </summary>
public class CoreDefinitions
{
    public const string InstrOnExcel = "OnExcel";
    public const string InstrOnSheet = "OnSheet";
    public const string InstrFirstRow = "FirstRow";
    public const string InstrForEach = "ForEach";

    public const string InstrRow = "Row";

    public const string InstrForEachRow = "ForEachRow";

    public const string InstrNext = "Next";

    public const string InstrCol = "Col";
    public const string InstrCell = "Cell";
    public const string InstrBgColor = "BgColor";
    public const string InstrFgColor = "FgColor";
    

    public const string InstrIf = "If";
    public const string InstrThen = "Then";
    public const string InstrElse = "Else";

    public const string InstrBlank = "Blank";
    public const string InstrNull = "Null";
    public const string InstrEndName = "End";

    public const string InstrAnd = "And";
    public const string InstrOr = "Or";

    public const string InstrDot = ".";
    public const string InstrComma = ",";
    

    //--function names
    public const string InstrFuncSelectFiles = "SelectFiles";
    public const string InstrCreateExcel = "CreateExcel";
    public const string InstrCopyHeader = "CopyHeader";
    public const string InstrCopyRow = "CopyRow";

    


    public const string InstrFuncDate = "Date";
    public const string InstrFuncDateTime = "DateTime";
    public const string InstrFuncTime = "Time";

    public static string DefaultExcelSheetName = "Sheet1";

    /// <summary>
    /// Used in OnExcel instr.
    /// The first data row to process.
    /// (after the header row)
    ///
    /// value in base1 (human readable).
    /// </summary>
    public static int FirstDataRowIndex = 2;

    //-system/builtin var
    public static string SysVarDateFormatName = "$DateFormat";

    public static string SysVarDateTimeFormatName = "$DateTimeFormat";

    public static string SysVarTimeFormatName = "$TimeFormat";

    //--$CurrencyFormat
    public static string SysVarCurrencyFormatName = "$CurrencyFormat";

    public static string SysVarStrCompareCaseSensitive = "$StrCompareCaseSensitive";

}