using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.System;
/// <summary>
/// An excel file.
/// Has sheets. One or more.
/// </summary>
public interface IExcelFile
{
    /// <summary>
    /// source file name.
    /// </summary>
    public string FileName { get; set; }

    // public bool IsOpen()
}
