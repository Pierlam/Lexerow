using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DevApp;
internal class DateSandBox
{
    public static void TestDates()
    {
        string fileName = @"Input\TestDate.xlsx";

        var stream = new FileStream(fileName, FileMode.Open);
        stream.Position = 0;
        var workbook = new XSSFWorkbook(stream);

        var sheet = workbook.GetSheetAt(0);

        //--B2 date short	11/01/2025
        var row = sheet.GetRow(1);
        var cell = row.GetCell(1);
        var cellType = cell.CellType;
        double valDouble = cell.NumericCellValue;

        // 14 -> "m/d/yy"  -> le format est faux!!! inversion jour et mois!!
        short dataFormat = cell.CellStyle.DataFormat;
        string format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        // {11/01/2025 00:00:00}
        DateTime dt = DateTime.FromOADate(valDouble);
        // {11/01/2025}
        DateOnly dateOnly = DateOnly.FromDateTime(dt);

        //--B3 date long: vendredi 23 février 2024
        row = sheet.GetRow(2);
        cell = row.GetCell(1);
        cellType = cell.CellType;
        valDouble = cell.NumericCellValue;

        // 165 -> "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy"
        dataFormat = cell.CellStyle.DataFormat;
        format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);
        // {23/02/2024 00:00:00}
        dt = DateTime.FromOADate(valDouble);

        //--B4 hour:	11:20:45
        row = sheet.GetRow(3);
        cell = row.GetCell(1);
        cellType = cell.CellType;
        valDouble = cell.NumericCellValue;

        // 166 -> "[$-F400]h:mm:ss\\ AM/PM"
        dataFormat = cell.CellStyle.DataFormat;
        format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        // {30/12/1899 11:20:45}
        dt = DateTime.FromOADate(valDouble);

        //--B5 dateTime	4/10/23 11:34:56
        row = sheet.GetRow(4);
        cell = row.GetCell(1);
        cellType = cell.CellType;
        valDouble = cell.NumericCellValue;

        // 169 -> "d/m/yy\\ h:mm;@"
        dataFormat = cell.CellStyle.DataFormat;
        format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        // {04/10/2023 11:34:56}  mois et jour à l'anglaise
        dt = DateTime.FromOADate(valDouble);

        //--close
        using var writeStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);
        workbook.Write(writeStream);
        workbook.Close();
        stream.Close();
    }
}
