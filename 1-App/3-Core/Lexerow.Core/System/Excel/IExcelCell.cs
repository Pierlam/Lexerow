using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// cell raw value type.
/// voir CellType
/// </summary>
public enum CellRawValueType
{
    Unknow,
    String,
    Numeric,
    DateTime,
    DateOnly,
    TimeOnly,

    // cell exists but the content is empty
    Blank,

    // TODO: gérer +tard
    //Formula,
}

public interface IExcelCell
{
    int RowNum { get; set; }
    int ColNum { get; set; }

    double GetRawValueNumeric();

    string GetRawValueString();
}
