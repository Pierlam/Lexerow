using Lexerow.Core.ExcelLayer;
using Lexerow.Core.System;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace DevApp;
internal class DateSandBox
{
    /// <summary>
    /// write different date format.
    /// </summary>
    public static void TestWriteDate()
    {
        string fileName = @"Input\TestWriteDates.xlsx";

        var stream = new FileStream(fileName, FileMode.Open);
        stream.Position = 0;
        var workbook = new XSSFWorkbook(stream);

        var sheet = workbook.GetSheetAt(0);

        //--B2 contains a string, the modifcation to a dateOnly type fails!
        var row = sheet.GetRow(1);
        var cell = row.GetCell(1);
        var cellType = cell.CellType;
        //double valDouble = cell.NumericCellValue;

        //ExcelCellNpoi excelCellNpoi = excelCell as ExcelCellNpoi;

        // force the type: 14 -> "m/d/yy" by default
        // TODO: PROBLEME!! all get the style !!!
        //cell.CellStyle.DataFormat = 14;

        DateOnly dateOnly = new DateOnly(2025, 12, 6);
        cell.SetCellValue(dateOnly);

        //--D2 is 02/10/2021 -> 6/11/2024
        // if it's already a dateOnly, the modification of the value works well
        row = sheet.GetRow(1);
        cell = row.GetCell(3);
        dateOnly = new DateOnly(2024, 11, 6);
        cell.SetCellValue(dateOnly);

        //--E2 is null -> 7/12/2023  no, doesn't work
        row = sheet.GetRow(1);
        cell = row.GetCell(4);
        cell= row.CreateCell(4);
        dateOnly = new DateOnly(2023, 12, 7);
        cell.SetCellValue(dateOnly);

        //--
        ICellStyle style = workbook.CreateCellStyle();
        IDataFormat dataFormat = workbook.CreateDataFormat();
        // 2019-07-15 00:00:00
        //style.DataFormat = dataFormat.GetFormat("yyyy-MM-dd HH:mm:ss");
        // 15/07/2019 
        style.DataFormat = dataFormat.GetFormat("dd/MM/yyyy");

        //--E3 Create a custom date format  OK!!
        row = sheet.GetRow(2);
        //cell = row.GetCell(4);
        cell = row.CreateCell(4);
        // Apply the style to the cell
        cell.CellStyle = style;
        dateOnly = new DateOnly(2019, 7, 15);
        cell.SetCellValue(dateOnly);


        //--E4 Create a custom date format  OK!!
        row = sheet.GetRow(3);
        cell = row.CreateCell(4);
        // Apply the style to the cell
        cell.CellStyle = style;
        dateOnly = new DateOnly(2018, 6, 14);
        cell.SetCellValue(dateOnly);


        //--close
        using var writeStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);
        workbook.Write(writeStream);
        workbook.Close();
        stream.Close();

    }

    /// <summary>
    /// read different date format.
    /// </summary>
    public static void TestReadDates()
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
