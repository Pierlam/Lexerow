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
    public List<ICellStyle> ListCellStyle { get; set; } = new List<ICellStyle>();

    public string FileName { get; set; }

    /// <summary>
    /// choix de le passer en public pour que le processor dédié y accède normalement
    /// (pas en private)
    /// </summary>
    public FileStream Stream { get; set; }

    public XSSFWorkbook XssWorkbook { get; }

    public ICellStyle GetStyle(CellRawValueType type, string dateFormat)
    {
        foreach (ICellStyle style in ListCellStyle) 
        {
            //if(style.DataFormat)
        }
        return null;
    }


    /// <summary>
    ///  Create a new style+Format and save it.
    /// </summary>
    /// <param name="cellInit"></param>
    /// <param name="type"></param>
    /// <param name="format"></param>
    /// <returns></returns>
    public ICellStyle CreateStyle(ICell cellInit, CellRawValueType type, string format)
    {
        var style = XssWorkbook.CreateCellStyle();
        // clone the style to keep it
        style.CloneStyleFrom(cellInit.CellStyle);
        IDataFormat dataFormat = XssWorkbook.CreateDataFormat();
        style.DataFormat = dataFormat.GetFormat(format);

        // save the new style
        // TODO:

        return style;
    }
}