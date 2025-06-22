using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

/// <summary>
/// Execution result insights/Informations.
/// </summary>
public class ExecResultInsights
{
    /// <summary>
    /// All Analyzed datarow counter in ForEach Row/If condition.
    /// </summary>
    public int AnalyzedDatarowCount { get; set; } = 0;

    /// <summary>
    /// all If condition matching in ForEach Row/If condition.
    /// </summary>
    public int IfCondMatchCount { get; set; } = 0;

    public void Clear()
    {
        IfCondMatchCount = 0;
        AnalyzedDatarowCount = 0;
    }

}
