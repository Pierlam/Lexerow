using Lexerow.Core.System;
using NPOI.HSSF.Util;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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
        if (Cell.CellType == CellType.Formula)
            return CellRawValueType.Formula;
        if (Cell.CellType == CellType.Blank)
            return CellRawValueType.Blank;

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

    public int GetBgColor()
    {
        // BUG: return 0 if nothing or always 64 if a color is set!!
        XSSFCellStyle style = (XSSFCellStyle)Cell.CellStyle;
        XSSFColor color= style.FillBackgroundXSSFColor;
        
        if (color == null) return 0;
        return color.Indexed;
    }
    public int GetFgColor()
    {
        // BUG: return 0 if nothing or always 64 if a color is set!!
        XSSFCellStyle style = (XSSFCellStyle)Cell.CellStyle;
        XSSFColor color = style.FillForegroundXSSFColor;

        if (color == null) return 0;
        return color.Indexed;
    }
}