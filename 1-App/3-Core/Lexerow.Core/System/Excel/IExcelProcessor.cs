using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.Excel;
public interface IExcelProcessor
{
    bool Open(string fileName, out IExcelFile excelFile, out CoreError coreError);

    bool Close(IExcelFile excelFile);

    IExcelSheet GetSheetAt(IExcelFile excelFile, int index);

    IExcelRow GetRowAt(IExcelSheet sheet, int rowNum);

    IExcelCell GetCellAt(IExcelSheet excelSheet, int rowNum, int colNum);

    CellRawValueType GetCellValueType(IExcelSheet excelSheet, IExcelCell excelCell);

    int GetLastRowNum(IExcelSheet excelSheet);

    public IExcelCell CreateCell(IExcelSheet excelSheet, int rowNum, int colNum);

    // create a cell
    bool CreateCell(IExcelSheet excelSheet, int rowNum, int colNum, string value);

    bool CreateCell(IExcelSheet excelSheet, int rowNum, int colNum, int value);

    bool CreateCell(IExcelSheet excelSheet, int rowNum, int colNum, double value);

    bool CreateCell(IExcelSheet excelSheet, int rowNum, int colNum, DateTime value);

    bool DeleteCell(IExcelSheet excelSheet, int rowNum, int colNum);

    // set the new value to the cell
    bool SetCellValue(IExcelCell excelCell, double value);

    bool SetCellValue(IExcelCell excelCell, string value);

    bool SetCellValueString(IExcelCell cell, string value);

    bool SetCellValueInt(IExcelCell cell, int value);

    bool SetCellValueDouble(IExcelCell cell, double value);

    bool SetCellValueDateOnly(IExcelCell excelCell, ValueDateOnly value);

    bool SetCellValueDateTime(IExcelCell excelCell, ValueDateTime value);

    bool SetCellValueTimeOnly(IExcelCell excelCell, ValueTimeOnly value);

    bool SetCellValueBlank(IExcelCell excelCell);

    bool Save(IExcelFile excelFile);

}
