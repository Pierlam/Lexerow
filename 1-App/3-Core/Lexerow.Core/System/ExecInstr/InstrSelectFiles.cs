using Lexerow.Core.System.Compilator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;

public enum InstrSelectFilesSelector
{
    Select,
    Unselect
}

/// <summary>
/// One final filename, with full path.
/// The instr source is attached, exp: *.xslx
/// InstrSelectedFilename
/// </summary>
public class InstrSelectedFilename
{
    public InstrSelectedFilename(InstrBase instrBase, string filename)
    {
        InstrBase = instrBase;
        Filename = filename;
    }

    /// <summary>
    /// instr source, exp: *.xlsx
    /// </summary>
    public InstrBase InstrBase { get; set; }

    /// <summary>
    /// the selected filename, matching the instruction.
    /// With the full path.
    /// </summary>
    public string Filename { get; set; }

}

/// <summary>
/// instruction: SelectFiles("myfile.xlsx")
/// select filenames to be exact.
/// </summary>
public class InstrSelectFiles : InstrBase
{
    public InstrSelectFiles(ScriptToken scriptToken):base(scriptToken)
    {
        InstrType = InstrType.SelectFiles;
        IsFunctionCall = true;
        // return a list of filename
        ReturnType = InstrFunctionReturnType.ListString;
    }

    /// <summary>
    /// List of parameters of the function.
    /// exp: "data.xlsx", "file.xlsx", "*.xlsx", filename
    /// constValue/string or varname, or fctcall or string concatenation expression.
    /// </summary>
    public List<InstrBase> ListInstrParams { get; private set; }=new List<InstrBase>();

    /// <summary>
    /// list of FileSelector (+ or -), exactly one for each parameter. The first one is always +/Select
    /// </summary>
    public List<InstrSelectFilesSelector> ListFilesSelectors { get; private set; } =new List<InstrSelectFilesSelector>();

    /// <summary>
    /// used to execute the instruction.
    /// Index set before the first one.
    /// </summary>
    public int CurrParamNum { get; set; } = -1;

    /// <summary>
    /// List of final parameters, after processing complex parameters (fct call, string concatenation)
    /// contains only const value, filename string (no var)
    /// Used in execution/Run of the instr.
    /// </summary>
    public List<InstrBase> RunTmpListFinalInstrParams { get; private set; } = new List<InstrBase>();

    /// <summary>
    /// List of final filename obtained after decoding parameters of the function SelectFiles.
    /// exp: *.xlsx can match several files.
    /// </summary>
    public List<InstrSelectedFilename> ListSelectedFilename { get; private set; }= new List<InstrSelectedFilename>();

    public void AddParamSelect(InstrBase param)
    {
        ListInstrParams.Add(param);
        ListFilesSelectors.Add(InstrSelectFilesSelector.Select);
    }
    public void AddParamUnSelect(InstrBase param)
    {
        ListInstrParams.Add(param);
        ListFilesSelectors.Add(InstrSelectFilesSelector.Unselect);
    }

    public void AddFinalFilename(InstrBase instrSource, string filename)
    {
        InstrSelectedFilename instrSelectFilesFilename = new InstrSelectedFilename(instrSource, filename);
        ListSelectedFilename.Add(instrSelectFilesFilename);
    }
}
