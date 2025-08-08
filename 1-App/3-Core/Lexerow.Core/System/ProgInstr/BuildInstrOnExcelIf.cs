using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Build instruction If used in OnExcel OnSheet instruction.
/// exp: If A.Cell > 12
/// </summary>
public class BuildInstrOnExcelIf
{
    public BuildInstrOnExcelIf(ExecTokExcelCellAddress excelCellAddress, ExecTokExcelFuncCell execTokExcelFuncCell, ExecTokCompOperator execTokCompOperator, InstrConstValue execTokConstValue)
    {
        ExecTokExcelCellAddress= excelCellAddress;
        ExecTokExcelFuncCell= execTokExcelFuncCell;
        ExecTokCompOperator= execTokCompOperator;
        ExecTokConstValue = execTokConstValue;
    }

    public ExecTokExcelCellAddress ExecTokExcelCellAddress { get; set; }
    public ExecTokExcelFuncCell ExecTokExcelFuncCell { get; set; }
    public  ExecTokCompOperator ExecTokCompOperator { get; set; }
    public InstrConstValue ExecTokConstValue { get; set; }
}
