﻿using Lexerow.Core.System;
using Lexerow.Core.Tests._20_Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Tests.Exec;

/// <summary>
/// test Execute InstrForEachIfThen on DateTime type.
/// </summary>
[TestClass]
public class ExecInstrForEachRowIfThenDateTimeTests
{
    /// <summary>
    /// col:    C/2           D/3
    /// row1: 10/02/2019   10/2/19 11:05
    /// row2: 20/09/2023   20/9/23 12:45
    /// 
    /// Col C -> DateOnly
    /// Col D -> DateTime
    /// </summary>
    [TestMethod]
    public void TestBasicDateOnly()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\ForIfDateOnlyBasic.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--Create: col C, c.Value > 01/02/2020
        InstrCompCellVal exprCompIf = core.Builder.CreateInstrCompCellVal(2, InstrCompCellValOperator.GreaterThan, new DateOnly(2020,02,01));

        //--Create: col C, c.Value:=01/02/2020
        InstrSetCellVal instrSetVal = core.Builder.CreateInstrSetCellVal(2, new DateOnly(2020, 02, 01));

        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, exprCompIf, instrSetVal);
        Assert.IsTrue(execResult.Result);


        execResult = core.Exec.Compile();
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
        Assert.IsTrue(execResult.Result);

        //--check the content of modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);

        //--check, row2, colC/2: 01/02/2020
        bool res = ExcelChecker.CheckCellValue(wb, 0, 2, 2, new DateOnly(2020, 02, 01));
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }

    /// <summary>
    /// col:    C/2           D/3
    /// row1: 10/02/2019   10/2/19 11:05
    /// row2: 20/09/2023   20/9/23 12:45
    /// 
    /// Col C -> DateOnly
    /// Col D -> DateTime
    /// </summary>
    [TestMethod]
    public void TestBasicDateTime()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\ForIfDateTimeBasic.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--Create: col D, cell.Value > 01/02/2020
        InstrCompCellVal exprCompIf = core.Builder.CreateInstrCompCellVal(3, InstrCompCellValOperator.GreaterThan, new DateTime(2020, 02, 01));

        //--Create: col D, c.Value:=01/02/2020 09:27:59
        InstrSetCellVal instrSetVal = core.Builder.CreateInstrSetCellVal(3, new DateTime(2020, 02, 01,9,27,59));

        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, exprCompIf, instrSetVal);
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Compile();
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
        Assert.IsTrue(execResult.Result);

        //--check the content of modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);

        //--check, row2, colD/3: 01/02/2020
        bool res = ExcelChecker.CheckCellValue(wb, 0, 2, 3, new DateTime(2020, 02, 01,9,27,59));
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }

    /// <summary>
    /// row        Col C
    ///  1   16:27:34	Perso
    ///  2   09:43:56	Perso   -> modified
    ///  3   12:45:00	Heure
    ///  4  08:24:15	Heure   -> modified
    /// </summary>
    [TestMethod]
    public void TestBasicTimeOnly()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\ForIfTimeOnlyBasic.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--Create: col C/2, cell.Value < 10:20:30
        InstrCompCellVal exprCompIf = core.Builder.CreateInstrCompCellVal(2, InstrCompCellValOperator.LesserThan, new TimeOnly(10,20,30));

        //--Create: col C, c.Value:= 10:20:30
        InstrSetCellVal instrSetVal = core.Builder.CreateInstrSetCellVal(2, new TimeOnly(10,20,30));

        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, exprCompIf, instrSetVal);
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Compile();
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
        Assert.IsTrue(execResult.Result);

        //--check the content of modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);

        //--check, row2, colC/2: not modified
        bool res = ExcelChecker.CheckCellValue(wb, 0, 1, 2, new TimeOnly(16, 27, 34));
        Assert.IsTrue(res);
        // modified
        res = ExcelChecker.CheckCellValue(wb, 0, 2, 2, new TimeOnly(10,20,30));
        Assert.IsTrue(res);

        // modified
        res = ExcelChecker.CheckCellValue(wb, 0, 4, 2, new TimeOnly(10, 20, 30));
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }

    /// <summary>
    /// On col(A) If Cell.Value > '01/01/2020' Then Cell.Value= '01/01/2020'
    /// If cell is DateTime, apply it.
    /// If Cell is TimeOnly, not applicable
    /// 
    /// 0        col D/3
    /// 1   datetime 10/2/19 11:05      -> 01/02/2020 00:00
    /// 2   datetime 20/9/23 12:45
    /// </summary>
    [TestMethod]
    public void IfThenDateOnlyOnDateTimeOk()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\ForIfThenDateOnlyOnDateTime.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--Create: col D, cell.Value < 01/02/2020
        InstrCompCellVal exprCompIf = core.Builder.CreateInstrCompCellVal(3, InstrCompCellValOperator.LesserThan, new DateOnly(2020, 02, 01));

        //--Create: col D, c.Value := 01/02/2020
        InstrSetCellVal instrSetVal = core.Builder.CreateInstrSetCellVal(3, new DateOnly(2020, 02, 01));

        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, exprCompIf, instrSetVal);
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Compile();
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
        Assert.IsTrue(execResult.Result);

        //--check the content of modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);

        //--check, row2, colD/3: 01/02/2020
        bool res = ExcelChecker.CheckCellValue(wb, 0, 1, 3, new DateTime(2020, 02, 01, 0, 0, 0));
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }

    /// <summary>
    ///  0   date
    ///  1   10/02/2019  -> 01/02/2020
    ///  2   05/11/2023
    /// </summary>
    [TestMethod]
    public void IfThenDateTimeOnDateOnlyOk()
    {
        LexerowCore core = new LexerowCore();

        string fileName = @"10-Files\ForIfThenDateTimeOnDateOnly.xlsx";

        ExecResult execResult = core.Builder.CreateInstrOpenExcel("file", fileName);

        //--Create: col D, cell.Value < 01/02/2020  12:34:56
        InstrCompCellVal exprCompIf = core.Builder.CreateInstrCompCellVal(3, InstrCompCellValOperator.LesserThan, new DateTime(2020, 02, 01, 12,34,56));

        //--Create: col D, c.Value := 01/02/2020 12:34:56
        InstrSetCellVal instrSetVal = core.Builder.CreateInstrSetCellVal(3, new DateTime(2020, 02, 01, 12, 34, 56));

        execResult = core.Builder.CreateInstrForEachRowIfThen("file", 0, 1, exprCompIf, instrSetVal);
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Compile();
        Assert.IsTrue(execResult.Result);

        execResult = core.Exec.Execute();
        Assert.IsTrue(execResult.Result);

        //--check the content of modified excel file
        var stream = ExcelChecker.OpenExcel(fileName);
        var wb = ExcelChecker.GetWorkbook(stream);

        //--check, row2, colD/3: 01/02/2020, without time!!
        bool res = ExcelChecker.CheckCellValue(wb, 0, 1, 3, new DateTime(2020, 02, 01, 0, 0, 0));
        Assert.IsTrue(res);

        ExcelChecker.CloseExcel(stream);
    }
}
