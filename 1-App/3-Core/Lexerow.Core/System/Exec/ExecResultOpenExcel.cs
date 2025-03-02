using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

// TODO: supprimer
public class ExecResultOpenExcel_DEL
{
    public CoreError Error { get; set; } = null;

    public IExcelFile? ExcelFile { get; set; } = null;
}
