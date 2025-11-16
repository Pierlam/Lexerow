namespace Lexerow.Core.System;

public class FileSheetProcessed
{
    public FileSheetProcessed(string currFilename)
    {
        CurrFilename = currFilename;
    }

    public string CurrFilename { get; set; } = string.Empty;

    public int SheetNum { get; set; } = -1;

    public int RowCount { get; set; } = 0;

    /// <summary>
    /// all If condition matching in ForEach Row/If condition.
    /// </summary>
    public int IfCondMatchCount { get; set; } = 0;
}

/// <summary>
/// Execution result insights/Informations.
/// </summary>
public class ExecResultInsights
{
    public List<FileSheetProcessed> ListFileSheetProcessed { get; set; } = new List<FileSheetProcessed>();

    /// <summary>
    ///  Total count of files processed
    /// </summary>
    public int FileTotalCount { get; set; } = 0;

    /// <summary>
    ///  Total count of sheet processed
    /// </summary>
    public int SheetTotalCount { get; set; } = 0;

    /// <summary>
    /// All Analyzed datarow counter in ForEach Row/If condition.
    /// </summary>
    public int RowTotalCount { get; set; } = 0;

    /// <summary>
    /// all If condition matching in ForEach Row/If condition.
    /// </summary>
    public int IfCondMatchTotalCount { get; set; } = 0;

    /// <summary>
    /// The current object, during the execution of a program.
    /// </summary>
    public FileSheetProcessed CurrFileSheetProcessed { get; set; } = null;

    public void Clear()
    {
        FileTotalCount = 0;
        SheetTotalCount = 0;
        IfCondMatchTotalCount = 0;
        RowTotalCount = 0;
    }

    public void StartNewFile(string fileName)
    {
        CurrFileSheetProcessed = new FileSheetProcessed(fileName);
        ListFileSheetProcessed.Add(CurrFileSheetProcessed);
        FileTotalCount++;
    }

    public void StartNewSheet(int sheetNum)
    {
        if (CurrFileSheetProcessed == null) return;
        CurrFileSheetProcessed.SheetNum = sheetNum;
        SheetTotalCount++;
    }

    public void NewRowProcessed()
    {
        if (CurrFileSheetProcessed == null) return;
        CurrFileSheetProcessed.RowCount++;
        RowTotalCount++;
    }

    public void NewIfCondMatch()
    {
        if (CurrFileSheetProcessed == null) return;
        CurrFileSheetProcessed.IfCondMatchCount++;
        IfCondMatchTotalCount++;
    }
}