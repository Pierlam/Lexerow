using Lexerow.Core.System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;

namespace Lexerow.Core.ExcelLayer;

public class ExcelFileNpoi : IExcelFile
{
    public ExcelFileNpoi(string fileName, FileStream stream)
    {
        Stream = stream;
        XssWorkbook = new XSSFWorkbook(stream);
        FileName = fileName;
    }
    public List<CellStyleNpoi> ListCellStyle { get; set; } = new List<CellStyleNpoi>();

    public string FileName { get; set; }

    /// <summary>
    /// choix de le passer en public pour que le processor dédié y accède normalement
    /// (pas en private)
    /// </summary>
    public FileStream Stream { get; set; }

    public XSSFWorkbook XssWorkbook { get; }

    public XSSFCellStyle GetStyle(CellRawValueType type, string dataFormat, int bgColor, int fgColor)
    {
        foreach (CellStyleNpoi cellStyle in ListCellStyle) 
        {
            if (cellStyle.ValueType == type && cellStyle.DataFormat.Equals(dataFormat) && cellStyle.BgColor==bgColor && cellStyle.FgColor==fgColor)
                return cellStyle.CellStyle;
        }
        return null;
    }


    /// <summary>
    ///  Create a new style+Format and save it.
    /// </summary>
    /// <param name="cellInit"></param>
    /// <param name="type"></param>
    /// <param name="DataFormat"></param>
    /// <returns></returns>
    public XSSFCellStyle CreateStyle(ICell cellInit, CellRawValueType valueType, string DataFormat, int bgColor, int fgColor)
    {
        XSSFCellStyle style = (XSSFCellStyle)XssWorkbook.CreateCellStyle();
        // clone the style to keep it
        style.CloneStyleFrom(cellInit.CellStyle);
        IDataFormat dataFormat = XssWorkbook.CreateDataFormat();
        style.DataFormat = dataFormat.GetFormat(DataFormat);

        // save the new style
        CellStyleNpoi cellStyleNpoi= new CellStyleNpoi(valueType, DataFormat, style, bgColor, fgColor);
        ListCellStyle.Add(cellStyleNpoi);
        return style;
    }
}