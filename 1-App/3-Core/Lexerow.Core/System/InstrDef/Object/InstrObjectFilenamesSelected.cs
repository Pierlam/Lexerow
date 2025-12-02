using Lexerow.Core.System.InstrDef.Func;
using Lexerow.Core.System.ScriptDef;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System.InstrDef.Object;
public class InstrObjectFilenamesSelected : InstrBase
{
    public InstrObjectFilenamesSelected(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.ObjectFilenamesSelected;
    }
    /// <summary>
    /// List of final filename obtained after decoding parameters of the function SelectFiles.
    /// exp: *.xlsx can match several files.
    /// </summary>
    public List<InstrFuncSelectedFilename> ListSelectedFilename { get; private set; } = new List<InstrFuncSelectedFilename>();

    public void AddFinalFilename(InstrBase instrSource, string filename)
    {
        InstrFuncSelectedFilename instrSelectFilesFilename = new InstrFuncSelectedFilename(instrSource, filename);
        ListSelectedFilename.Add(instrSelectFilesFilename);
    }

}
