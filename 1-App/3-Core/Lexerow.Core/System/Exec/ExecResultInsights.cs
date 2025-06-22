using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

public class ExecResultInsights
{
    public int IfCondMatchCount { get; set; } = 0;

    public int AnalyzedDatarowCount { get; set; } = 0;

    //public int DatarowModifiedCount { get; set; } = 0;

    //public int DatarowCreatedCount { get; set; } = 0;

    //public int DatarowRemovedCount { get; set; } = 0;

    public void Clear()
    {
        IfCondMatchCount = 0;
        AnalyzedDatarowCount = 0;
    }

}
