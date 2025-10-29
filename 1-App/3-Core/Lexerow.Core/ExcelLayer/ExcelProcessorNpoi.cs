using Lexerow.Core.System.Excel;
using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.XWPF.UserModel;

namespace Lexerow.Core.ExcelLayer;

/// <summary>
/// Excel processor, based on NPOI implementation.
/// </summary>
public class ExcelProcessorNpoi : IExcelProcessor
{
    public bool Open(string fileName, out IExcelFile excelFile, out ExecResultError error)
    {
        if (!File.Exists(fileName)) 
        {
            excelFile = null;
            error = new ExecResultError(ErrorCode.FileNotFound, fileName);
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
            error= new ExecResultError(ErrorCode.ExcelUnableOpenFile, ex.Message);
            return false;
        }
    }
    public bool Close(IExcelFile excelFile, out ExecResultError error)
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
            error = new ExecResultError(ErrorCode.ExcelUnableCloseFile, ex.Message);
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
        var row= excelSheetNpoi.Sheet.GetRow(rowNum);
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

        var row= excelSheetNpoi.Sheet.GetRow(rowNum);
        if (row == null) return null;
        var cell= row.GetCell(colNum);
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
        ExcelSheetNpoi excelSheetNpoi=  excelSheet as ExcelSheetNpoi;
        ExcelCellNpoi excelCellNpoi= excelCell as ExcelCellNpoi;

        if (excelCellNpoi.Cell.CellType == CellType.Blank)
            return CellRawValueType.Blank;

        // type string
        if (excelCellNpoi.Cell.CellType== CellType.String)
            return CellRawValueType.String;

        if (excelCellNpoi.Cell.CellType == CellType.Numeric)
        {
            // can be a int/double or date type
            // 0 -> 'General'
            short dataFormat = excelCellNpoi.Cell.CellStyle.DataFormat;
            string format = excelSheetNpoi.Sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

            // contains a date ?
            bool hasDate = false;
            if(format.Contains("yy") || format.Contains("?/?"))
                hasDate=true;

            // contains a time ?
            bool hasTime = false;
            if (format.Contains("h") || format.Contains("ss"))
                hasTime=true;

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
    /// The row should exists.
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

        var cell = row.CreateCell(colNum);
        IExcelCell excelCell= new ExcelCellNpoi(cell);
        return excelCell;
    }

    /// <summary>
    /// The row should exists.
    /// </summary>
    /// <param name="excelSheet"></param>
    /// <param name="rowNum"></param>
    /// <param name="colNum"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public bool CreateCell(IExcelSheet excelSheet, int rowNum, int colNum, string value)
    {
        ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;
        var row = excelSheetNpoi.Sheet.GetRow(rowNum);
        if (row == null) return false;

        var cell= row.CreateCell(colNum);
        cell.SetCellValue(value);
        return true;
    }

    public bool CreateCell(IExcelSheet excelSheet, int rowNum, int colNum, int value)
    {
        ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;
        var row = excelSheetNpoi.Sheet.GetRow(rowNum);
        if (row == null) return false;

        var cell = row.CreateCell(colNum);
        cell.SetCellValue(value);
        return true;
    }

    public bool CreateCell(IExcelSheet excelSheet, int rowNum, int colNum, double value)
    {
        ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;
        var row = excelSheetNpoi.Sheet.GetRow(rowNum);
        if (row == null) return false;

        var cell = row.CreateCell(colNum);
        cell.SetCellValue(value);
        return true;
    }

    public bool CreateCell(IExcelSheet excelSheet, int rowNum, int colNum, DateTime value)
    {
        ExcelSheetNpoi excelSheetNpoi = excelSheet as ExcelSheetNpoi;
        var row = excelSheetNpoi.Sheet.GetRow(rowNum);
        if (row == null) return false;

        var cell = row.CreateCell(colNum);
        cell.SetCellValue(value);
        return true;
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

    // set the new value to the cell
    public bool SetCellValue(IExcelCell excelCell, double value)
    {
        ExcelCellNpoi excelCellNpoi= excelCell as ExcelCellNpoi;

        excelCellNpoi.Cell.SetCellValue(value);
        return true;
    }

    public bool SetCellValue(IExcelCell excelCell, string value)
    {
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;

        excelCellNpoi.Cell.SetCellValue(value);
        return true;
    }

    public bool SetCellValueString(IExcelCell excelCell, string value)
    {
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;

        // force the string type: 0 'General'
        excelCellNpoi.Cell.CellStyle.DataFormat = 0;
        excelCellNpoi.Cell.SetCellValue(value);
        return true;

    }

    public bool SetCellValueInt(IExcelCell excelCell, int value)
    {
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;

        // force the int type: 1 '0'
        excelCellNpoi.Cell.CellStyle.DataFormat = 1;
        excelCellNpoi.Cell.SetCellValue(value);
        return true;

    }

    public bool SetCellValueDouble(IExcelCell excelCell, double value)
    {
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;

        // force the double type: 2 '0.00'
        excelCellNpoi.Cell.CellStyle.DataFormat = 2;
        excelCellNpoi.Cell.SetCellValue(value);
        return true;

    }

    public bool SetCellValueDateOnly(IExcelCell excelCell, ValueDateOnly value)
    {
        ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;

        // force the type: 14 -> "m/d/yy" by default
        excelCellNpoi.Cell.CellStyle.DataFormat = 14;
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
