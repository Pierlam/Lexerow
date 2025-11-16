using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public interface IExcelSheet
{
    int Index { get; }
    string Name { get; }
    IExcelFile ExcelFileParent { get; }

}
