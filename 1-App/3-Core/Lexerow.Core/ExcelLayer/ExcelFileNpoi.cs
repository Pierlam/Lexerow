using Lexerow.Core.System;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.ExcelLayer;
public class ExcelFileNpoi : IExcelFile
{
    public ExcelFileNpoi(string fileName, FileStream stream)
    {
        Stream = stream;
        XssWorkbook = new XSSFWorkbook(stream);
        FileName = fileName;
    }

    public string FileName { get; set; }


    /// <summary>
    /// choix de le passer en public pour que le processor dédié y accède normalement
    /// (pas en private)
    /// </summary>
    public FileStream Stream { get; set; }

    public XSSFWorkbook XssWorkbook { get; }

}
