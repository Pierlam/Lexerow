using Lexerow.Core.System.Excel;
using NPOI.SS.UserModel;

namespace Lexerow.Core.ExcelLayer;

public class ExcelRowNpoi : IExcelRow
{
    public ExcelRowNpoi(IRow row)
    {
        Row = row;
    }

    public IRow Row { get; private set; }
}