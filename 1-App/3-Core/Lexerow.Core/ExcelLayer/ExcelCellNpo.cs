using Lexerow.Core.System;
using NPOI.SS.UserModel;

namespace Lexerow.Core.ExcelLayer;

public class ExcelCellNpoi : IExcelCell
{
    public ExcelCellNpoi(ICell cell)
    {
        Cell = cell;
    }

    public ICell Cell { get; set; }

    public int RowNum
    { get { return Cell.RowIndex; } set { } }
    public int ColNum
    { get { return Cell.ColumnIndex; } set { } }

    public CellRawValueType GetRawValueType()
    {
        if (Cell == null)
            return CellRawValueType.Unknow;

        if (Cell.CellType == CellType.String)
            return CellRawValueType.String;
        if (Cell.CellType == CellType.Numeric)
            return CellRawValueType.Numeric;
        //if (Cell.CellType == CellType.Formula)
        //    return CellRawValueType.Formula;
        //if (Cell.CellType == CellType.Blank)
        //    return CellRawValueType.Blank;

        return CellRawValueType.Unknow;
    }

    public double GetRawValueNumeric()
    {
        if (Cell.CellType != CellType.Numeric) return 0.0;

        return Cell.NumericCellValue;
    }

    public string GetRawValueString()
    {
        if (Cell.CellType == CellType.String)
            return Cell.StringCellValue;

        if (Cell.CellType == CellType.Numeric)
            return Cell.NumericCellValue.ToString();

        return string.Empty;
    }
}