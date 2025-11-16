using Lexerow.Core.Utils;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests._20_Utils;

/*
public class ExcelFileSheetTest
{
    public string FileName {  get; set; }  

    public XSSFWorkbook Workbook { get; set; }

    public int SheetNul {  get; set; }  
}*/

public class ExcelTestChecker
{
    public static FileStream OpenExcel(string fileName)
    {
        var stream = new FileStream(fileName, FileMode.Open);
        stream.Position = 0;
        return stream;
    }

    public static XSSFWorkbook GetWorkbook(FileStream stream)
    {
        return new XSSFWorkbook(stream);
    }

    public static void  CloseExcel(FileStream stream)
    {
        stream.Close();
    }

    public static bool CheckCellValueColRow(XSSFWorkbook workbook, int sheetNum, int colNum, int rowNum, double expectedValue)
    {
        return CheckCellValue(workbook, sheetNum, rowNum, colNum, expectedValue);
    }

    public static bool CheckCellValueColRow(XSSFWorkbook workbook, int sheetNum, int colNum, int rowNum, string expectedValue)
    {
        return CheckCellValue(workbook, sheetNum, rowNum, colNum, expectedValue);
    }

    public static bool CheckCellValue(XSSFWorkbook workbook, int sheetNum, string cell, double expectedValue)
    {
        if(!ExcelExtendedUtils.SplitCellAddress(cell, out string colName, out int colIndex, out int rowIndex))
            return false;
        return CheckCellValue(workbook,sheetNum,rowIndex-1,colIndex-1, expectedValue);
    }

    public static bool CheckCellValue(XSSFWorkbook workbook, int sheetNum, int rowNum, int colNum, double expectedValue)
    {
        var sheet = workbook.GetSheetAt(sheetNum);
        var row = sheet.GetRow(rowNum);
        var cell = row.GetCell(colNum);
        if (cell==null)return false;

        if (cell.CellType != CellType.Numeric) return false;
        if (cell.NumericCellValue == null) return false;

        double val = cell.NumericCellValue;

        return expectedValue== val;
    }

    /// <summary>
    /// cell=A2
    /// </summary>
    /// <param name="workbook"></param>
    /// <param name="sheetNum"></param>
    /// <param name="cell"></param>
    /// <param name="expectedValue"></param>
    /// <returns></returns>
    public static bool CheckCellValue(XSSFWorkbook workbook, int sheetNum, string cell, string expectedValue)
    {
        if (!ExcelExtendedUtils.SplitCellAddress(cell, out string colName, out int colIndex, out int rowIndex))
            return false;
        return CheckCellValue(workbook, sheetNum, rowIndex - 1, colIndex - 1, expectedValue);
    }

    public static bool CheckCellValue(XSSFWorkbook workbook, int sheetNum, int rowNum, int colNum, string expectedValue)
    {
        var sheet = workbook.GetSheetAt(sheetNum);
        var row = sheet.GetRow(rowNum);
        var cell = row.GetCell(colNum);
        if (cell == null) return false;
        if (cell.CellType != CellType.String)  return false;

        string val = cell.StringCellValue;

        return expectedValue == val;
    }

    public static bool CheckCellValue(XSSFWorkbook workbook, int sheetNum, int rowNum, int colNum, DateOnly expectedValue)
    {
        var sheet = workbook.GetSheetAt(sheetNum);
        var row = sheet.GetRow(rowNum);
        var cell = row.GetCell(colNum);
        if (cell == null) return false;

        if (cell.CellType != CellType.Numeric) return false;

        double val = cell.NumericCellValue;

        DateTime expectedDateTime = expectedValue.ToDateTime(TimeOnly.MinValue);

        return expectedDateTime == DateTime.FromOADate(val);
    }

    public static bool CheckCellValue(XSSFWorkbook workbook, int sheetNum, int rowNum, int colNum, DateTime expectedValue)
    {
        var sheet = workbook.GetSheetAt(sheetNum);
        var row = sheet.GetRow(rowNum);
        var cell = row.GetCell(colNum);
        if (cell == null) return false;

        if (cell.CellType != CellType.Numeric) return false;

        double val = cell.NumericCellValue;

        return expectedValue == DateTime.FromOADate(val);
    }

    public static bool CheckCellValue(XSSFWorkbook workbook, int sheetNum, int rowNum, int colNum, TimeOnly expectedValue)
    {
        var sheet = workbook.GetSheetAt(sheetNum);
        var row = sheet.GetRow(rowNum);
        var cell = row.GetCell(colNum);
        if (cell == null) return false;

        if (cell.CellType != CellType.Numeric) return false;

        double val = cell.NumericCellValue;
        TimeOnly foundVal= DateTimeUtils.ToTimeOnly(val);

        return foundVal == expectedValue;
    }

    public static bool CheckCellNull(XSSFWorkbook workbook, int sheetNum, int rowNum, int colNum)
    {
        var sheet = workbook.GetSheetAt(sheetNum);
        var row = sheet.GetRow(rowNum);
        var cell = row.GetCell(colNum);
        // ok
        if (cell == null) return true;

        return false;
    }

    public static bool CheckCellValueBlank(XSSFWorkbook workbook, int sheetNum, string cell)
    {
        if (!ExcelExtendedUtils.SplitCellAddress(cell, out string colName, out int colIndex, out int rowIndex))
            return false;
        return CheckCellValueBlank(workbook, sheetNum, rowIndex-1, colIndex - 1);
    }

    public static bool CheckCellValueBlank(XSSFWorkbook workbook, int sheetNum, int rowNum, int colNum)
    {
        var sheet = workbook.GetSheetAt(sheetNum);
        var row = sheet.GetRow(rowNum);
        var cell = row.GetCell(colNum);
        if (cell == null) return false;

        if (cell.CellType == CellType.Blank) 
            // ok
            return true;


        return false;
    }

}
