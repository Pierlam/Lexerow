using NPOI.OpenXmlFormats.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Instruction comparison for cell value.
/// exp: A.Cell In ["y", "ok" ]
/// By default, the comparison is case sensitive.
///
/// exp: A.Cell In /I ["y", "ok" ]  -> comparison is case insensitive
/// </summary>
public class InstrCompColCellInStringItems : InstrRetBoolBase
{
    public InstrCompColCellInStringItems(int colNum, bool inOrNotIn, bool caseSensitive, List<string> listItems)
    {
        ColNum = colNum;
        InOrNotIn = inOrNotIn;
        CaseSensitive = caseSensitive;
        ListItems = listItems;
    }

    public int ColNum { get; set; }

    public bool InOrNotIn { get; set; }
    public bool CaseSensitive { get; set; }

    public List<string> ListItems { get; private set; }
}
