using Lexerow.Core.System.ScriptDef;

namespace Lexerow.Core.System.InstrDef.Object;

/// <summary>
/// One final filename, with full path.
/// The instr source is attached, exp: *.xslx
/// </summary>
public class ObjectSelectedFile
{
    public ObjectSelectedFile(InstrBase instrBase, string filename)
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
/// Object selected files.
/// Result of the execution of the instr: InstrSelectFiles
/// InstrObjectSelectedFiles
/// </summary>
public class InstrObjectSelectedFiles : InstrBase
{
    public InstrObjectSelectedFiles(ScriptToken scriptToken) : base(scriptToken)
    {
        InstrType = InstrType.ObjectSelectedFiles;
    }

    /// <summary>
    /// List of final filename obtained after decoding parameters of the function SelectFiles.
    /// exp: *.xlsx can match several files.
    /// </summary>
    public List<ObjectSelectedFile> ListObjectSelectedFile { get; private set; } = new List<ObjectSelectedFile>();

    public void AddFinalFilename(InstrBase instrSource, string filename)
    {
        ObjectSelectedFile objectSelectedFile = new ObjectSelectedFile(instrSource, filename);
        ListObjectSelectedFile.Add(objectSelectedFile);
    }
}