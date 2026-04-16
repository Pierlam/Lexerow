using Lexerow.Core.System.ActivLog;
using Lexerow.Core.System.InstrDef;
using System;

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

    /// <summary>
    /// Peek/read the top item on the stack.
    /// </summary>
    /// <returns></returns>
    public InstrBase Peek()
    {
        var instr = StackInstr.Peek();
        Log(_logger, "CompilStackInstr.Peek", instr);
        return instr;
    }

    /// <summary>
    /// Pop/Remove the top item on the stack.
    /// </summary>
    /// <returns></returns>
    public InstrBase Pop()
    {
        try
        {
            var instr = StackInstr.Pop();
            Log(_logger, "CompilStackInstr.Pop", instr);
            return instr;
        }
        catch (Exception e)
        {
            _logger.LogCompilError("CompilStackInstr.Pop", e.ToString());
            throw e;
        }
    }

    public void Push(InstrBase instr)
    {
        StackInstr.Push(instr);
        Log(_logger, "CompilStackInstr.Push", instr);
    }

    /// <summary>
    /// Get/read the inst just before the instr on top of the stack.
    /// </summary>
    /// <param name="stkInstr"></param>
    /// <returns></returns>
    public InstrBase ReadInstrBeforeTop()
    {
        // need to have 2 instr on the stack
        if (StackInstr.Count < 2) return null;
        return StackInstr.ElementAt(1);
    }

    public InstrBase ReadInstrBeforeBeforeTop()
    {
        // need to have 3 instr on the stack
        if (StackInstr.Count < 3) return null;
        return StackInstr.ElementAt(2);
    }

    /// <summary>
    /// Looking for an instr in the stack, starting from the top.
    /// </summary>
    /// <param name="stkInstr"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    public InstrBase FindFirstInstrFromTop(InstrType type, InstrType type2)
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

    /// <summary>
    /// Looking for an instr in the stack, starting from the top (latest instr added).
    /// </summary>
    /// <param name="stkInstr"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    public InstrBase FindInstrFromTop(InstrType type)
    {
        foreach (var instr in StackInstr)
        {
            if (instr.InstrType == type)
                return instr;
        }
        return null;
    }

    /// <summary>
    /// Looking for an instr in the stack, starting from the top (latest instr added).
    /// if the instr is not found, return 0.
    /// </summary>
    /// <param name="stkInstr"></param>
    /// <param name="type"></param>
    /// <returns></returns>
    public int GetDistanceFromTop(InstrType type)
    {
        int i = 1;
        foreach (var instr in StackInstr)
        {
            if (instr.InstrType == type)
                return i;
            i++;
        }
        // instr not found
        return 0;
    }

    /// <summary>
    /// Save instr in a list of instr in the right order.
    /// </summary>
    /// <param name="instrCount"></param>
    /// <returns></returns>
    public List<InstrBase> SaveInListReverse(int instrCount)
    {
        int i = 1;
        List<InstrBase> list = new List<InstrBase>();
        foreach (var instr in StackInstr)
        {
            if (i > StackInstr.Count) break;
            if (i > instrCount) break;
            list.Add(instr);
            i++;
        }
        list.Reverse();
        return list;
    }

    /// <summary>
    /// Save instr in a list of instr in the right order.
    /// </summary>
    /// <param name="instrCount"></param>
    /// <returns></returns>
    public List<InstrBase> RemoveSaveInListReverse(int instrCount)
    {
        int i = 1;
        List<InstrBase> list = new List<InstrBase>();
        while(true)
        {
            // no more item in the stack, bye
            if (StackInstr.Count==0) break;

            if (i > instrCount) break;

            var instr = StackInstr.Pop();
            list.Add(instr);
            i++;
        }
        list.Reverse();
        return list;
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

    void Log(IActivityLogger logger, string msg , InstrBase instrBase)
    {
        int count = StackInstr.Count;
        if(logger.IsLevelTraceActive())
            logger.LogCompil(ActivityLogLevel.Trace, msg, instrBase.ToString() + ", Nb=" + count.ToString() + " Top> " + Dump()) ;
    }

}