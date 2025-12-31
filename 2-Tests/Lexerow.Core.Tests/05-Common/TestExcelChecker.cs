using DocumentFormat.OpenXml.Spreadsheet;
using Lexerow.Core.Utils;
using Microsoft.Testing.Platform.OutputDevice;
using OpenExcelSdk;

namespace Lexerow.Core.Tests._20_Utils;


public class TestExcelChecker
{
    static OpenExcelSdk.ExcelProcessor excelProcessor= new OpenExcelSdk.ExcelProcessor();
    static StyleMgr styleMgr= new StyleMgr();


    public static ExcelFile Open(string filename)
    {
        ExcelFile excelFile= excelProcessor.OpenExcelFile(filename);
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
    /// <param name="cellReference"></param>
    /// <param name="numFormatId"></param>
    /// <returns></returns>
    public static bool CheckCellStyleDataFormat(ExcelFile excelFile, string cellReference,  int numFormatId)
    {
        ExcelSheet excelSheet = excelProcessor.GetSheetAt(excelFile, 0);

        // get the cell at the address
        ExcelCell excelCell = excelProcessor.GetCellAt(excelSheet, cellReference);
        if (excelCell == null) return false;

        var cellFormat = excelProcessor.GetCellFormat(excelSheet, excelCell);

        // no format
        if(cellFormat == null) return false;
        if (cellFormat.ApplyNumberFormat == null) return false;
        return (numFormatId == (int)cellFormat.NumberFormatId.Value);
    }

    /// <summary>
    /// Check the number format  of the value of the cell.
    /// Should be a custom one, Id>163.
    /// </summary>
    /// <param name="excelFile"></param>
    /// <param name="cellReference"></param>
    /// <param name="numFormat"></param>
    /// <returns></returns>
    public static bool CheckCellStyleDataFormat(ExcelFile excelFile, string cellReference,  string numFormat)
    {
        ExcelSheet excelSheet = excelProcessor.GetSheetAt(excelFile, 0);
            return false;

        // get the cell at the address
        ExcelCell excelCell = excelProcessor.GetCellAt(excelSheet, cellReference);
            return false;
        if (excelCell == null) return false;

        var cellFormat = excelProcessor.GetCellFormat(excelSheet, excelCell);

        if (cellFormat.ApplyNumberFormat == null) return false;
        
        styleMgr.GetCustomNumberFormat(excelSheet, cellFormat.NumberFormatId.Value, out string numFormatGet);
        return (numFormatGet == numFormat);
    }


    public static bool CheckCellBgColor(ExcelFile excelFile, string cellReference, int colorCode)
    {
        ExcelSheet excelSheet = excelProcessor.GetSheetAt(excelFile, 0);
            return false;

        // get the cell at the address
        ExcelCell excelCell = excelProcessor.GetCellAt(excelSheet, cellReference);
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
    public static bool CheckCellValue(ExcelFile excelFile, string cellReference, double expectedValue)
    {
        // get the first sheet
        ExcelSheet excelSheet = excelProcessor.GetSheetAt(excelFile, 0);
            return false;

        // get the cell at the address
        ExcelCell excelCell = excelProcessor.GetCellAt(excelSheet, cellReference);
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        ExcelCellValue excelCellValue = excelProcessor.GetCellValue(excelSheet, excelCell);
            return false;

        // can be a double value
        if(excelCellValue.CellType == ExcelCellType.Double)
        {
            // compare values
            if (expectedValue != excelCellValue.DoubleValue) return false;
            return true;
        }

        // can be an int value
        if (excelCellValue.CellType == ExcelCellType.Integer)
        {
            // compare values
            if (expectedValue != excelCellValue.IntegerValue) return false;
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
    public static bool CheckCellValue(ExcelFile excelFile, string cellReference, string expectedValue)
    {
        // get the first sheet
        ExcelSheet excelSheet = excelProcessor.GetSheetAt(excelFile, 0);
            return false;

        // get the cell at the address
        ExcelCell excelCell = excelProcessor.GetCellAt(excelSheet, cellReference);
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        ExcelCellValue excelCellValue = excelProcessor.GetCellValue(excelSheet, excelCell);
            return false;

        // type string expected
        if (excelCellValue.CellType != ExcelCellType.String) return false;

        // compare values
        if (expectedValue != excelCellValue.StringValue) return false;

        return true;
    }

    /// <summary>
    /// Get the cell type and value of a cell, in the first sheet by default.
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="cell"></param>
    /// <param name="expectedValue"></param>
    /// <returns></returns>
    public static bool CheckCellValue(ExcelFile excelFile, string cellReference, DateOnly expectedValue)
    {
        // get the first sheet
        ExcelSheet excelSheet = excelProcessor.GetSheetAt(excelFile, 0);
            return false;

        // get the cell at the address
        ExcelCell excelCell = excelProcessor.GetCellAt(excelSheet, cellReference);
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        ExcelCellValue excelCellValue = excelProcessor.GetCellValue(excelSheet, excelCell);
            return false;

        // type dateOnly expected
        if (excelCellValue.CellType != ExcelCellType.DateOnly) return false;

        // compare values
        if (expectedValue != excelCellValue.DateOnlyValue) return false;

        return true;
    }

    /// <summary>
    /// Get the cell type and value of a cell, in the first sheet by default.
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="cell"></param>
    /// <param name="expectedValue"></param>
    /// <returns></returns>
    public static bool CheckCellValue(ExcelFile excelFile, string cellReference, DateTime expectedValue)
    {
        // get the first sheet
        ExcelSheet excelSheet = excelProcessor.GetSheetAt(excelFile, 0);
            return false;

        // get the cell at the address
        ExcelCell excelCell = excelProcessor.GetCellAt(excelSheet, cellReference);
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        ExcelCellValue excelCellValue = excelProcessor.GetCellValue(excelSheet, excelCell);
            return false;

        // type dateTime expected
        if (excelCellValue.CellType != ExcelCellType.DateTime) return false;

        // compare values
        if (expectedValue != excelCellValue.DateTimeValue) return false;

        return true;
    }

    /// <summary>
    /// Get the cell type and value of a cell, in the first sheet by default.
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="cell"></param>
    /// <param name="expectedValue"></param>
    /// <returns></returns>
    public static bool CheckCellValue(ExcelFile excelFile, string cellReference, TimeOnly expectedValue)
    {
        // get the first sheet
        ExcelSheet excelSheet = excelProcessor.GetSheetAt(excelFile, 0);
            return false;

        // get the cell at the address
        ExcelCell excelCell = excelProcessor.GetCellAt(excelSheet, cellReference);
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        ExcelCellValue excelCellValue = excelProcessor.GetCellValue(excelSheet, excelCell);
            return false;

        // type TimeOnly expected
        if (excelCellValue.CellType != ExcelCellType.TimeOnly) return false;

        // compare values
        if (expectedValue != excelCellValue.TimeOnlyValue) return false;

        return true;
    }

    /// <summary>
    /// Get the cell type and value of a cell, in the first sheet by default.
    /// </summary>
    /// <param name="proc"></param>
    /// <param name="cell"></param>
    /// <param name="expectedValue"></param>
    /// <returns></returns>
    public static bool CheckCellValueEmpty(ExcelFile excelFile, string cellReference)
    {
        // get the first sheet
        ExcelSheet excelSheet = excelProcessor.GetSheetAt(excelFile, 0);
            return false;

        // get the cell at the address
        ExcelCell excelCell = excelProcessor.GetCellAt(excelSheet, cellReference);
            return false;
        if (excelCell == null) return false;

        // get the type and the value of the cell
        ExcelCellValue excelCellValue = excelProcessor.GetCellValue(excelSheet, excelCell);
            return false;

        // can be type string
        if (excelCellValue.CellType == ExcelCellType.String)
        {
            if(excelCellValue.StringValue==string.Empty)return true;
            return false;
        }

        // type undefined expected
        if (excelCellValue.CellType != ExcelCellType.Undefined) return false;

        return true;
    }

}