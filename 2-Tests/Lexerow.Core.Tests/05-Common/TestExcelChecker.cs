using DocumentFormat.OpenXml.Spreadsheet;
using Lexerow.Core.Utils;
using Microsoft.Testing.Platform.OutputDevice;
using OpenExcelSdk;
using OpenExcelSdk.System;

namespace Lexerow.Core.Tests._20_Utils;


public class TestExcelChecker
{
    static OpenExcelSdk.ExcelProcessor excelProcessor= new OpenExcelSdk.ExcelProcessor();
    static StyleMgr styleMgr= new StyleMgr();


    public static ExcelFile Open(string filename)
    {
        bool res = excelProcessor.Open(filename, out ExcelFile excelFile, out var error);
        if (!res) return null;

        return excelFile;
    }

    public static FileStream OpenExcel(string fileName)
    {
        var stream = new FileStream(fileName, FileMode.Open);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Check the number format Id of the value of the cell.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="addressName"></param>
    /// <param name="numFormatId"></param>
    /// <returns></returns>
    public static bool CheckCellStyleDataFormat(ExcelFile excelFile, string addressName,  int numFormatId)
    {
        if (!excelProcessor.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out ExcelError error))
            return false;

        // get the cell at the address
        if (!excelProcessor.GetCellAt(excelSheet, addressName, out ExcelCell excelCell, out ExcelError excelError))
            return false;
        if (excelCell == null) return false;

        var cellFormat = excelProcessor.GetCellFormat(excelSheet, excelCell);

        if (cellFormat.ApplyNumberFormat == null) return false;
        return (numFormatId == (int)cellFormat.NumberFormatId.Value);
    }

    /// <summary>
    /// Check the number format  of the value of the cell.
    /// Should be a custom one, Id>163.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="addressName"></param>
    /// <param name="numFormat"></param>
    /// <returns></returns>
    public static bool CheckCellStyleDataFormat(ExcelFile excelFile, string addressName,  string numFormat)
    {
        if (!excelProcessor.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out ExcelError error))
            return false;

        // get the cell at the address
        if (!excelProcessor.GetCellAt(excelSheet, addressName, out ExcelCell excelCell, out ExcelError excelError))
            return false;
        if (excelCell == null) return false;

        var cellFormat = excelProcessor.GetCellFormat(excelSheet, excelCell);

        if (cellFormat.ApplyNumberFormat == null) return false;
        
        styleMgr.GetCustomNumberFormat(excelSheet, cellFormat.NumberFormatId.Value, out string numFormatGet);
        return (numFormatGet == numFormat);
    }


    public static bool CheckCellBgColor(ExcelFile excelFile, string addressName, int colorCode)
    {
        if (!excelProcessor.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out ExcelError error))
            return false;

        // get the cell at the address
        if (!excelProcessor.GetCellAt(excelSheet, addressName, out ExcelCell excelCell, out ExcelError excelError))
            return false;
        if (excelCell == null) return false;

        var cellFormat = excelProcessor.GetCellFormat(excelSheet, excelCell);

        if (cellFormat.ApplyFill== null) return false;

        //cellFormat.FillId

        //XX
        //ICellStyle cellStyle = cell.CellStyle;

        //IColor col = cellStyle.FillBackgroundColorColor;
        //if (col == null) return false;

        //if (col.Indexed != colorCode) return false;
        //return true;

        return false;
    }


    /// <summary>
    /// Get the cell type and value of a cell, in the first sheet by default.
    /// exp: A3= 12.5 ?
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="cell"></param>
    /// <param name="expectedValue"></param>
    /// <returns></returns>
    public static bool CheckCellValue(ExcelFile excelFile, string addressName, double expectedValue)
    {
        // get the first sheet
        if(!excelProcessor.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out ExcelError error))
            return false;

        // get the cell at the address
        if (!excelProcessor.GetCellAt(excelSheet, addressName, out ExcelCell excelCell, out ExcelError excelError))
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        if(!excelProcessor.GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out excelError))
            return false;

        // can be a double value
        if(excelCellValueMulti.CellType == ExcelCellType.Double)
        {
            // compare values
            if (expectedValue != excelCellValueMulti.DoubleValue) return false;
            return true;
        }

        // can be an int value
        if (excelCellValueMulti.CellType == ExcelCellType.Integer)
        {
            // compare values
            if (expectedValue != excelCellValueMulti.IntegerValue) return false;
            return true;
        }


        return true;
    }

    /// <summary>
    /// Get the cell type and value of a cell, in the first sheet by default.
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="cell"></param>
    /// <param name="expectedValue"></param>
    /// <returns></returns>
    public static bool CheckCellValue(ExcelFile excelFile, string addressName, string expectedValue)
    {
        // get the first sheet
        if (!excelProcessor.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out ExcelError error))
            return false;

        // get the cell at the address
        if (!excelProcessor.GetCellAt(excelSheet, addressName, out ExcelCell excelCell, out ExcelError excelError))
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        if (!excelProcessor.GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out excelError))
            return false;

        // type string expected
        if (excelCellValueMulti.CellType != ExcelCellType.String) return false;

        // compare values
        if (expectedValue != excelCellValueMulti.StringValue) return false;

        return true;
    }

    /// <summary>
    /// Get the cell type and value of a cell, in the first sheet by default.
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="cell"></param>
    /// <param name="expectedValue"></param>
    /// <returns></returns>
    public static bool CheckCellValue(ExcelFile excelFile, string addressName, DateOnly expectedValue)
    {
        // get the first sheet
        if (!excelProcessor.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out ExcelError error))
            return false;

        // get the cell at the address
        if (!excelProcessor.GetCellAt(excelSheet, addressName, out ExcelCell excelCell, out ExcelError excelError))
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        if (!excelProcessor.GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out excelError))
            return false;

        // type dateOnly expected
        if (excelCellValueMulti.CellType != ExcelCellType.DateOnly) return false;

        // compare values
        if (expectedValue != excelCellValueMulti.DateOnlyValue) return false;

        return true;
    }

    /// <summary>
    /// Get the cell type and value of a cell, in the first sheet by default.
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="cell"></param>
    /// <param name="expectedValue"></param>
    /// <returns></returns>
    public static bool CheckCellValue(ExcelFile excelFile, string addressName, DateTime expectedValue)
    {
        // get the first sheet
        if (!excelProcessor.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out ExcelError error))
            return false;

        // get the cell at the address
        if (!excelProcessor.GetCellAt(excelSheet, addressName, out ExcelCell excelCell, out ExcelError excelError))
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        if (!excelProcessor.GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out excelError))
            return false;

        // type dateTime expected
        if (excelCellValueMulti.CellType != ExcelCellType.DateTime) return false;

        // compare values
        if (expectedValue != excelCellValueMulti.DateTimeValue) return false;

        return true;
    }

    /// <summary>
    /// Get the cell type and value of a cell, in the first sheet by default.
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="cell"></param>
    /// <param name="expectedValue"></param>
    /// <returns></returns>
    public static bool CheckCellValue(ExcelFile excelFile, string addressName, TimeOnly expectedValue)
    {
        // get the first sheet
        if (!excelProcessor.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out ExcelError error))
            return false;

        // get the cell at the address
        if (!excelProcessor.GetCellAt(excelSheet, addressName, out ExcelCell excelCell, out ExcelError excelError))
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        if (!excelProcessor.GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out excelError))
            return false;

        // type TimeOnly expected
        if (excelCellValueMulti.CellType != ExcelCellType.TimeOnly) return false;

        // compare values
        if (expectedValue != excelCellValueMulti.TimeOnlyValue) return false;

        return true;
    }

    /// <summary>
    /// Get the cell type and value of a cell, in the first sheet by default.
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="cell"></param>
    /// <param name="expectedValue"></param>
    /// <returns></returns>
    public static bool CheckCellValueEmpty(ExcelFile excelFile, string addressName)
    {
        // get the first sheet
        if (!excelProcessor.GetSheetAt(excelFile, 0, out ExcelSheet excelSheet, out ExcelError error))
            return false;

        // get the cell at the address
        if (!excelProcessor.GetCellAt(excelSheet, addressName, out ExcelCell excelCell, out ExcelError excelError))
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        if (!excelProcessor.GetCellTypeAndValue(excelSheet, excelCell, out ExcelCellValueMulti excelCellValueMulti, out excelError))
            return false;

        // type undefined expected
        if (excelCellValueMulti.CellType != ExcelCellType.Undefined) return false;

        return true;
    }

}