using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
public abstract class InstrCallMemberBase : InstrBase
{
    /// <summary>
    /// for InstrExcelFile, can be: FirstSheet, GetSheet[idx], GetSheet[name], ...
    /// </summary>
    public InstrCallMemberBase InstrCallMember { get; set; }
}
