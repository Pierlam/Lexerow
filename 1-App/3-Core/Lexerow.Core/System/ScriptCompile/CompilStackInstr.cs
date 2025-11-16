using Lexerow.Core.System.ActivLog;

namespace Lexerow.Core.System.ScriptCompile;

/// <summary>
/// Compilation Stack of instructions.
/// Used to parse script code.
/// </summary>
public class CompilStackInstr
{
    private IActivityLogger _logger;

    public CompilStackInstr(IActivityLogger logger)
    {
        _logger = logger;
    }

    public int Count
    { get { return StackInstr.Count; } }

    public InstrBase Peek()
    {
        var instr = StackInstr.Peek();
        _logger.LogCompilOnGoing(ActivityLogLevel.Detail, "CompilStackInstr.Peek", instr.ToString());
        return instr;
    }

    public InstrBase Pop()
    {
        try
        {
            var instr = StackInstr.Pop();
            _logger.LogCompilOnGoing(ActivityLogLevel.Detail, "CompilStackInstr.Pop", instr.ToString());
            return instr;
        }
        catch (Exception e)
        {
            _logger.LogCompilEndError(null, "CompilStackInstr.Pop", e.ToString());
            throw e;
        }
    }

    public void Push(InstrBase instr)
    {
        StackInstr.Push(instr);
    }

    /// <summary>
    /// Get the inst just before the instr on top of the stack.
    /// </summary>
    /// <param name="stkInstr"></param>
    /// <returns></returns>
    public InstrBase GetBeforeTop()
    {
        // need to ahve 2 isntr on the stack
        if (StackInstr.Count < 2) return null;
        return StackInstr.ElementAt(1);
    }

    /// <summary>
    /// Looking for an instr in the stack, starting from the top.
    /// </summary>
    /// <param name="stkInstr"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    public InstrBase FindFirstFromTop(InstrType type, InstrType type2)
    {
        foreach (var instr in StackInstr)
        {
            if (instr.InstrType == type)
                return instr;
            if (instr.InstrType == type2)
                return instr;
        }
        return null;
    }

    public string Dump()
    {
        string dump = string.Empty;
        foreach (var instr in StackInstr)
        {
            if (!string.IsNullOrEmpty(dump)) dump += " | ";
            dump += instr.ToString();
        }

        return dump;
    }

    /// <summary>
    /// the stack onf isntr, is private.
    /// </summary>
    private Stack<InstrBase> StackInstr { get; set; } = new Stack<InstrBase>();
}