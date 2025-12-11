using Lexerow.Core.ExcelLayer;
using Lexerow.Core.System;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DevApp;
public class DevExcelProcessorNpoi
{
    public static void Test()
    {
        ExcelProcessorNpoi npoi = new ExcelProcessorNpoi();

        string filename = @"Input\TestNpoi.xlsx";
        npoi.Open(filename, out IExcelFile excelFile, out ResultError error);

        var sheet= npoi.GetSheetAt(excelFile, 0);
        IExcelCell cell;

        //-B2: 12
        cell= npoi.GetCellAt(sheet, 1, 1);
        npoi.SetCellValue(cell, "coucou");

        //-E2
        cell = npoi.CreateCell(sheet, 1, 4);

        //-B3: hello
        cell = npoi.GetCellAt(sheet, 2, 1);
        npoi.SetCellValue(cell, 15);

        //-B4: 34,56
        cell = npoi.GetCellAt(sheet, 3, 1);
        npoi.SetCellValue(cell, "cool");

        //-B5: null
        cell = npoi.CreateCell(sheet, 4, 1);
        npoi.SetCellValue(cell, 56.78);

        //ICellStyle style = workbook.CreateCellStyle();
        //IDataFormat dataFormat = workbook.CreateDataFormat();
        //style.DataFormat = dataFormat.GetFormat("dd/MM/yyyy");

        //-B6: null
        cell = npoi.CreateCell(sheet, 5, 1);
        //npoi.SetCellValueDateOnly(cell, new ValueDateOnly(new DateOnly(2025,12,7)));
        npoi.SetCellValue(excelFile, sheet, cell, new DateOnly(2025, 12, 7), "dd/MM/yyyy");


        //-B7: int -> 13/10/2019
        cell = npoi.GetCellAt(sheet, 6, 1);
        //npoi.SetCellValueDateOnly(cell, new ValueDateOnly(new DateOnly(2019, 10, 13)));
        npoi.SetCellValue(excelFile, sheet, cell, new DateOnly(2019, 10, 13), "dd/MM/yyyy");

        //-B8: string -> 15/09/2018
        cell = npoi.GetCellAt(sheet, 7, 1);
        //npoi.SetCellValueDateOnly(cell, new ValueDateOnly(new DateOnly(2018, 09, 15)));
        npoi.SetCellValue(excelFile, sheet, cell, new DateOnly(2018, 09, 15), "dd/MM/yyyy");

        //-B9: blank -> 17/08/2017
        cell = npoi.GetCellAt(sheet, 8, 1);
        //npoi.SetCellValueDateOnly(cell, new ValueDateOnly(new DateOnly(2017, 08, 17)));
        npoi.SetCellValue(excelFile, sheet, cell, new DateOnly(2017, 08, 17), "dd/MM/yyyy");

        //-B10: date -> 18/04/1995
        cell = npoi.GetCellAt(sheet, 9, 1);
        //npoi.SetCellValueDateOnly(cell, new ValueDateOnly(new DateOnly(1995, 04, 18)));
        npoi.SetCellValue(excelFile, sheet, cell, new DateOnly(1995, 04, 18), "dd/MM/yyyy");

        //-B11: date -> 18
        cell = npoi.GetCellAt(sheet, 10, 1);
        npoi.SetCellValue(cell, 18);


        // save and close the excel file
        npoi.Save(excelFile);
        npoi.Close(excelFile, out error);
    }
}
