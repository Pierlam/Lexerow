using Lexerow.Core.ExcelLayer;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.Model;
using NPOI.XSSF.Streaming.Values;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DevApp;

/// <summary>
/// 
/// cell -> CellStyle.DataFormat
/// 
/// https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BuiltinFormats.html
/// </summary>
public class DevNpoi
{
    /// <summary>
    /// more cases are in this file:
    ///     Input\TestTypes.xlsx
    /// </summary>
    /// <param name="fileName"></param>
    public void TestBlankNull()
    {
        string fileName = @"Input\TestBlankNull.xlsx";

        var stream = new FileStream(fileName, FileMode.Open);
        stream.Position = 0;
        var workbook = new XSSFWorkbook(stream);

        var sheet = workbook.GetSheetAt(0);

        //--B2: blank
        var row = sheet.GetRow(1);
        var cell = row.GetCell(1);
        // blank
        var cellType = cell.CellType;
        string valStr = cell.StringCellValue;
        // 6 ->
        short dataFormat = cell.CellStyle.DataFormat;
        string format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        //--B3: null, cell removed by copy-paste
        row = sheet.GetRow(2);
        cell = row.GetCell(1);   // is null

        //--close
        using var writeStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);
        workbook.Write(writeStream);
        workbook.Close();
        stream.Close();

    }

    /// <summary>
    /// more cases are in this file:
    ///     Input\TestTypes.xlsx
    /// </summary>
    /// <param name="fileName"></param>
    public void TestTypes()
    {
        string fileName = @"Input\TestMoney.xlsx";

        var stream = new FileStream(fileName, FileMode.Open);
        stream.Position = 0;
        var workbook = new XSSFWorkbook(stream);

        var sheet = workbook.GetSheetAt(0);

        StylesTable stylesTable=  workbook.GetStylesSource();
        for (int i = 0; i <stylesTable.NumCellStyles; i++)
        {
            var style = stylesTable.GetStyleAt(i);
            // à chaque style est associé un DataFormat
        }

        // tables des data format
        // 165: "#,##0.00\\ \"€\""
        var numberFormats =stylesTable.GetNumberFormats();

        // force currency sur dollar 
        var styleCurrency = workbook.CreateCellStyle();
        // USD:  "[$$-409]#,##0.00"
        styleCurrency.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat("[$$-409]#,##0.00");

        // 5, "$#,##0_);($#,##0)"
        //string format1 = sheet.Workbook.CreateDataFormat().GetFormat(5);
        //styleCurrency.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat(format);
        //styleCurrency.DataFormat = 165;

        int r = 0;

        //--row0: standard: texte
        var row = sheet.GetRow(0);
        var cell = row.GetCell(1);
        var cellType = cell.CellType;
        string valStr = cell.StringCellValue;
        // 0 -> 'General'
        short dataFormat = cell.CellStyle.DataFormat;
        string format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        //--row1: nombre: 12
        row = sheet.GetRow(1);
        cell = row.GetCell(1);
        cellType = cell.CellType;
        double valNum = cell.NumericCellValue;
        // 1 -> '0'
        dataFormat = cell.CellStyle.DataFormat;
        format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        //--row2: nombre2: 23,45
        row = sheet.GetRow(2);
        cell = row.GetCell(1);
        cellType = cell.CellType;
        valNum = cell.NumericCellValue;
        // 2 -> '0.00'
        dataFormat = cell.CellStyle.DataFormat;
        format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        //--row3: monétaire: 56 €
        row = sheet.GetRow(3);
        cell = row.GetCell(1);
        cellType = cell.CellType;
        valNum = cell.NumericCellValue;
        cell.SetCellValue(57);
        // 164 -> "#,##0.00\\ \"€\""
        dataFormat = cell.CellStyle.DataFormat;
        format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        //--row4: Compta: 67 €
        row = sheet.GetRow(4);
        cell = row.GetCell(1);
        cellType = cell.CellType;
        valNum = cell.NumericCellValue;
        // 44 -> "_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_-"
        dataFormat = cell.CellStyle.DataFormat;
        format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        //--row5: Monetaire : $89,00
        row = sheet.GetRow(5);
        cell = row.GetCell(1);
        cellType = cell.CellType;
        valNum = cell.NumericCellValue;
        // 166 -> "[$$-409]#,##0.00"
        dataFormat = cell.CellStyle.DataFormat;
        format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        //--row6: Monétaire: 24 €
        row = sheet.GetRow(6);
        cell = row.GetCell(1);
        cellType = cell.CellType;
        valNum = cell.NumericCellValue;
        // 6 ->  "#,##0\\ \"€\";[Red]\\-#,##0\\ \"€\""
        dataFormat = cell.CellStyle.DataFormat;
        format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        //--row7: Monétaire: $36
        row = sheet.GetRow(7);
        cell = row.GetCell(1);
        cellType = cell.CellType;
        valNum = cell.NumericCellValue;
        // 167 ->  "[$$-409]#,##0"
        dataFormat = cell.CellStyle.DataFormat;
        format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);

        //--close
        using var writeStream = new FileStream(fileName, FileMode.Create, FileAccess.Write);
        workbook.Write(writeStream);
        workbook.Close();
        stream.Close();
    }

    public void TestDateTime(string fileName)
    {
        var stream = new FileStream(fileName, FileMode.Open);
        stream.Position = 0;
        var workbook = new XSSFWorkbook(stream);

        var sheet = workbook.GetSheetAt(0);

        //--B2 date short	11/01/2025
        var row = sheet.GetRow(1);
        var cell = row.GetCell(1);
        var cellType = cell.CellType;
        double valDouble = cell.NumericCellValue;
        short dataFormat = cell.CellStyle.DataFormat;
        // 14 -> "m/d/yy"  -> le format est faux!!! inversion jour et mois!!
        string format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);
        // {11/01/2025 00:00:00}
        DateTime dt = DateTime.FromOADate(valDouble);

        //--B3 date long: vendredi 23 février 2024
        row = sheet.GetRow(2);
        cell = row.GetCell(1);
        cellType = cell.CellType;
        valDouble = cell.NumericCellValue;
        // 165 -> "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy"
        dataFormat = cell.CellStyle.DataFormat;
        format = sheet.Workbook.CreateDataFormat().GetFormat(dataFormat);
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
        dataFormat = cell.CellStyle.DataFormat;
        // 169 -> "d/m/yy\\ h:mm;@"
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
