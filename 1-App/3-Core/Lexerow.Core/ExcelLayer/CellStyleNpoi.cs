using Lexerow.Core.System;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ExcelLayer;

/// <summary>
/// Represents an excel cell style.
/// </summary>
public class CellStyleNpoi
{
    public CellStyleNpoi(CellRawValueType valueType, string dataFormat, XSSFCellStyle cellStyle, int bgColor, int fgColor)
    {
        ValueType = valueType;
        DataFormat = dataFormat;
        CellStyle = cellStyle;
        BgColor = bgColor;
        FgColor = fgColor;
    }

    public CellRawValueType ValueType { get; set; }
    public string DataFormat { get; set; }
    public XSSFCellStyle CellStyle { get; set; }

    // background color
    public int BgColor { get; set; }

    // text/foreground color
    public int FgColor { get; set; }
}
