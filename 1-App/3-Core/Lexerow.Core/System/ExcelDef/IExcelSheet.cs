namespace Lexerow.Core.System;

public interface IExcelSheet
{
    int Index { get; }
    string Name { get; }
    IExcelFile ExcelFileParent { get; }
}