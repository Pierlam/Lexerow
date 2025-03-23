using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core;

public class CoreData
{
    /// <summary>
    /// Possible to create instructions only in build stage.
    /// </summary>
    public CoreStage Stage { get; set; } = CoreStage.Build;

    public List<InstrBase> ListInstr { get; set; } = new List<InstrBase>();
}
