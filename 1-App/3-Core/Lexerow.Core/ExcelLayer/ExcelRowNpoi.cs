using Lexerow.Core.System.Excel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ExcelLayer;
public class ExcelRowNpoi : IExcelRow
{
    public ExcelRowNpoi(IRow row)
    {
        Row = row;
    }

    public IRow Row { get; private set; }
}
