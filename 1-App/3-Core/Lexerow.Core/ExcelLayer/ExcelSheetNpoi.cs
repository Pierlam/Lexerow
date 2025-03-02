using Lexerow.Core.System;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ExcelLayer;
public class ExcelSheetNpoi : IExcelSheet
{
    public ExcelSheetNpoi(IExcelFile excelFileParent, int index, ISheet sheet)
    {
        ExcelFileParent = excelFileParent;
        Index = index;

        Sheet = sheet;

    }

    public IExcelFile ExcelFileParent { get; }
    public int Index { get; }

    public ISheet Sheet { get; }

    public string Name { get { return Sheet.SheetName; } }

}
