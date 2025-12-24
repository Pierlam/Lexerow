using Lexerow.Core.System;
using Lexerow.Core.System.Excel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace Lexerow.Core.ExcelLayer;

/// <summary>
/// Excel processor, based on NPOI implementation.
/// </summary>
public class ExcelProcessorNpoi : IExcelProcessor
{
    public bool Open(string fileName, out IExcelFile excelFile, out ResultError error)
    {
        if (!File.Exists(fileName))
        {
            excelFile = null;
            error = new ResultError(ErrorCode.FileNotFound, fileName);
            return false;
        }

        try
        {
            var stream = new FileStream(fileName, FileMode.Open);
            stream.Position = 0;

            excelFile = new ExcelFileNpoi(fileName, stream);
            error = null;
            return true;
        }
        catch (Exception ex)
        {
            excelFile = null;
            error = new ResultError(ErrorCode.ExcelUnableOpenFile, ex.Message);
            return false;
        }
    }

    public bool Close(IExcelFile excelFile, out ResultError error)
    {
        try
        {
            ExcelFileNpoi excelFileNpoi = excelFile as ExcelFileNpoi;

            excelFileNpoi.Stream.Close();
            error = null;
            return true;
        }
        catch (Exception ex)
        {
            excelFile = null;
            error = new ResultError(ErrorCode.ExcelUnableCloseFile, ex.Message);
            return false;
        }
    }

    public IExcelSheet GetSheetAt(IExcelFile excelFile, int index)
    {
        ExcelFileNpoi excelFileNpoi = excelFile as ExcelFileNpoi;

        if (index >= excelFileNpoi.XssWorkbook.NumberOfSheets) return null;

        var sheet = excelFileNpoi.XssWorkbook.GetSheetAt(index);

        ExcelSheetNpoi excelSheetNpoi = new ExcelSheetNpoi(excelFileNpoi, index, sheet);

        return excelSheetNpoi;
    }

    public IExcelRow GetRowAt(IExcelSheet excelSheet, int rowNum)
    {
        ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;
        var row = excelSheetNpoi.Sheet.GetRow(rowNum);
        if (row == null) return null;

        ExcelRowNpoi excelRowNpoi = new ExcelRowNpoi(row);
        return excelRowNpoi;
    }

    /// <summary>
    /// get the cell at the adress row and col num.
    /// base0
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="colNum"></param>
    /// <returns></returns>
    public IExcelCell GetCellAt(IExcelSheet excelSheet, int rowNum, int colNum)
    {
        ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;

        var row = excelSheetNpoi.Sheet.GetRow(rowNum);
        if (row == null) return null;
        var cell = row.GetCell(colNum);

        // no cell
        if (cell == null) return null;

        ExcelCellNpoi excelCellNpoi = new ExcelCellNpoi(cell);

        return excelCellNpoi;
    }

    /// <summary>
    /// Top, Left corner.
    /// </summary>
    /// <param name="excelSheet"></param>
    public void FindFirstCell(IExcelSheet excelSheet)
    {
        ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;

        //heet.FirstRowNum + 1); i <= sheet.LastRowNum;
        int rn = excelSheetNpoi.Sheet.FirstRowNum;
    }

    /// <summary>
    /// Return the raw cell value type.
    /// String, Numeric, DateTime, DateOnly, TimeOnly.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="excelCell"></param>
    /// <returns></returns>
    public CellRawValueType GetCellValueType(IExcelSheet excelSheet, IExcelCell excelCell)
    {
        ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;

        if (excelCellNpoi.Cell.CellType == CellType.Blank)
            return CellRawValueType.Blank;

        // type string
        if (excelCellNpoi.Cell.CellType == CellType.String)
            return CellRawValueType.String;

        if (excelCellNpoi.Cell.CellType == CellType.Numeric)
        {
            // can be a int/double or date type
            // 0 -> 'General'
            short dataFormat = excelCellNpoi.Cell.CellStyle.DataFormat;
            string format = excelSheetNpoi.Sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

            // contains a date ?
            bool hasDate = false;
            if (format.Contains("yy") || format.Contains("?/?"))
                hasDate = true;

            // contains a time ?
            bool hasTime = false;
            if (format.Contains("h") || format.Contains("ss"))
                hasTime = true;

            if (hasDate && hasTime)
                return CellRawValueType.DateTime;

            if (hasDate)
                return CellRawValueType.DateOnly;

            if (hasTime)
                return CellRawValueType.TimeOnly;

            // no date, no time
            return CellRawValueType.Numeric;
        }

        return CellRawValueType.Unknow;
    }

    /// <summary>
    /// return the last row number.
    /// 0-based
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <returns></returns>
    public int GetLastRowNum(IExcelSheet excelSheet)
    {
        ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;
        return excelSheetNpoi.Sheet.LastRowNum;
    }

    /// <summary>
    /// Create a new cell in a row.
    /// The row should exists.
    /// The type of the cell is blank.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="colNum"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public IExcelCell CreateCell(IExcelSheet excelSheet, int rowNum, int colNum)
    {
        ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;
        var row = excelSheetNpoi.Sheet.GetRow(rowNum);
        if (row == null) return null;

        // Type=blank, create a default style
        var cell = row.CreateCell(colNum);
        IExcelCell excelCell = new ExcelCellNpoi(cell);
        return excelCell;
    }

    public bool DeleteCell(IExcelSheet excelSheet, int rowNum, int colNum)
    {
        ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;
        var row = excelSheetNpoi.Sheet.GetRow(rowNum);
        if (row == null) return false;

        var cell = row.GetCell(colNum);
        if (cell == null) return false;

        row.RemoveCell(cell);
        return true;
    }

    /// <summary>
    /// Set the new double/int value to the cell.
    /// </summary>
    /// <param name="excelCell"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValue(IExcelCell excelCell, double value)
    {
        if (excelCell == null) return false;
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;
        if(excelCellNpoi.Cell==null) return false;

        // is it the cell contains a formula?
        if (excelCellNpoi.Cell.CellType == CellType.Formula)
            excelCellNpoi.Cell.SetCellType(CellType.Numeric);

        excelCellNpoi.Cell.SetCellValue(value);
        return true;
    }

    /// <summary>
    /// Set a string value to a cell.
    /// Replace the previous value/type.
    /// </summary>
    /// <param name="excelCell"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool SetCellValue(IExcelCell excelCell, string value)
    {
        if (excelCell == null) return false;
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;
        if (excelCellNpoi.Cell == null) return false;

        // is it the cell contains a formula?
        if (excelCellNpoi.Cell.CellType== CellType.Formula) 
            excelCellNpoi.Cell.SetCellType(CellType.String);

        excelCellNpoi.Cell.SetCellValue(value);
        return true;
    }

    /// <summary>
    /// Set a date value to a cell.
    /// Replace the previous value/type.
    /// More complex than basic cases.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="excelCell"></param>
    /// <param name="value"></param>
    /// <param name="format"></param>
    /// <returns></returns>
    public bool SetCellValue(IExcelFile excelFile, IExcelSheet excelSheet, IExcelCell excelCell, DateOnly value, string format)
    {
        if (excelCell == null) return false;
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;
        if (excelCellNpoi.Cell == null) return false;


        // get the type of the cell
        var cellType = GetCellValueType(excelSheet, excelCell);

        //--is date only, same format
        if (cellType == CellRawValueType.DateOnly)
        {
            short dataFormat = excelCellNpoi.Cell.CellStyle.DataFormat;
            ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;
            string currFormat = excelSheetNpoi.Sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);
            if(currFormat.Equals(format))
            {
                excelCellNpoi.Cell.SetCellValue(value);
                return true;
            }
        }

        //--is the style date-Format already defined?
        ExcelFileNpoi excelFileNpoi = excelFile as ExcelFileNpoi;
        int bgColor = excelCellNpoi.GetBgColor();
        int fgColor = excelCellNpoi.GetFgColor();
        ICellStyle style = excelFileNpoi.GetStyle(CellRawValueType.DateOnly, format, bgColor, fgColor);
        if(style == null) 
        {
            // create a new style+Format and save it            
            style= excelFileNpoi.CreateStyle(excelCellNpoi.Cell, CellRawValueType.DateOnly, format, bgColor, fgColor);
        }


        // is it the cell contains a formula?
        if (excelCellNpoi.Cell.CellType == CellType.Formula)
            excelCellNpoi.Cell.SetCellType(CellType.Numeric);

        // apply the new/existing style
        excelCellNpoi.Cell.CellStyle = style;

        excelCellNpoi.Cell.SetCellValue(value);
        return true;
    }

    public bool SetCellValue_BAK(IExcelFile excelFile, IExcelCell excelCell, DateOnly value, string format)
    {
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;
        if(excelCellNpoi.Cell == null) return false;

        ExcelFileNpoi excelFileNpoi= excelFile as ExcelFileNpoi;

        //-new style + new DataFormat
        ICellStyle style = excelFileNpoi.XssWorkbook.CreateCellStyle();
        IDataFormat dataFormat = excelFileNpoi.XssWorkbook.CreateDataFormat();
        style.DataFormat = dataFormat.GetFormat(format);
        excelCellNpoi.Cell.CellStyle = style;

        //-use existing style  -> doesn't work!
        //excelCellNpoi.Cell.CellStyle.DataFormat = dataFormat.GetFormat(format);

        DateTime val = value.ToDateTime(TimeOnly.MinValue);
        excelCellNpoi.Cell.SetCellValue(val.ToOADate());
        return true;
    }

    public bool SetCellValue_TEST(IExcelFile excelFile, IExcelCell excelCell, DateOnly value, string format)
    {
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;
        if (excelCellNpoi.Cell == null) return false;

        ExcelFileNpoi excelFileNpoi = excelFile as ExcelFileNpoi;

        //-new style
        ICellStyle style = excelFileNpoi.XssWorkbook.CreateCellStyle();
        // clone the style to keep it
        style.CloneStyleFrom(excelCellNpoi.Cell.CellStyle);

        //-new DataFormat
        IDataFormat dataFormat = excelFileNpoi.XssWorkbook.CreateDataFormat();
        style.DataFormat = dataFormat.GetFormat(format);
        excelCellNpoi.Cell.CellStyle = style;

        //-use existing style  -> doesn't work!
        //excelCellNpoi.Cell.CellStyle.DataFormat = dataFormat.GetFormat(format);

        DateTime val = value.ToDateTime(TimeOnly.MinValue);
        excelCellNpoi.Cell.SetCellValue(val.ToOADate());
        return true;
    }

    public bool SetCellValueDateOnly(IExcelCell excelCell, ValueDateOnly value)
    {
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;

        // force the type: 14 -> "m/d/yy" by default
        // TODO: PROBLEME!! all get the style !!!
        //excelCellNpoi.Cell.CellStyle.DataFormat = 14;
        excelCellNpoi.Cell.SetCellValue(value.ToDouble());
        return true;
    }

    public bool SetCellValueTimeOnly(IExcelCell excelCell, ValueTimeOnly value)
    {
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;

        // Force the type TimeOnly: 21/0x15, "h:mm:ss"
        excelCellNpoi.Cell.CellStyle.DataFormat = 21;
        excelCellNpoi.Cell.SetCellValue(value.ToDouble());
        return true;
    }

    public bool SetCellValueDateTime(IExcelCell excelCell, ValueDateTime value)
    {
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;

        // DateTime: 22/0x16, "m/d/yy h:mm"
        excelCellNpoi.Cell.CellStyle.DataFormat = 22;
        excelCellNpoi.Cell.SetCellValue(value.ToDouble());
        return true;
    }

    public bool SetCellValueBlank(IExcelCell excelCell)
    {
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;
        excelCellNpoi.Cell.SetBlank();
        return true;
    }

    public bool Save(IExcelFile excelFile)
    {
        ExcelFileNpoi excelFileNpoi = excelFile as ExcelFileNpoi;

        using var writeStream = new FileStream(excelFileNpoi.FileName, FileMode.Create, FileAccess.Write);
        excelFileNpoi.XssWorkbook.Write(writeStream);

        // TODO:
        excelFileNpoi.XssWorkbook.Close();
        return false;
    }
}