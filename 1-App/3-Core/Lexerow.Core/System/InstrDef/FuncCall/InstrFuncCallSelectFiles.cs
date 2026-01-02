using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef.FuncCall;

public enum InstrFuncSelectFilesSelector
{
    Select,
    Unselect
}

/// <summary>
/// instruction: SelectFiles("myfile.xlsx")
/// select filenames to be exact.
/// </summary>
public class InstrFuncCallSelectFiles : InstrBase
{
    public InstrFuncCallSelectFiles(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.FuncSelectFiles;
        IsFunctionCall = true;
        // return a list of filename
        ReturnType = InstrReturnType.ListString;
    }

    /// <summary>
    /// List of parameters of the function.
    /// exp: "data.xlsx", "file.xlsx", "*.xlsx", filename
    /// constValue/string or varname, or fctcall or string concatenation expression.
    /// </summary>
    public List<InstrBase> ListInstrParams { get; private set; } = new List<InstrBase>();

    /// <summary>
    /// list of FileSelector (+ or -), exactly one for each parameter. The first one is always +/Select
    /// </summary>
    public List<InstrFuncSelectFilesSelector> ListFilesSelectors { get; private set; } = new List<InstrFuncSelectFilesSelector>();

    /// <summary>
    /// used to execute the instruction.
    /// Index set before the first one.
    /// </summary>
    public int CurrParamNum { get; set; } = -1;

    public void AddParamSelect(InstrBase param)
    {
        ListInstrParams.Add(param);
        ListFilesSelectors.Add(InstrFuncSelectFilesSelector.Select);
    }

    public void AddParamUnSelect(InstrBase param)
    {
        ListInstrParams.Add(param);
        ListFilesSelectors.Add(InstrFuncSelectFilesSelector.Unselect);
    }
}