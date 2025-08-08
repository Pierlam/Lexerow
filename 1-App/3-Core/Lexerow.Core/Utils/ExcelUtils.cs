using Lexerow.Core.System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lexerow.Core.Utils;

/// <summary>
/// Excel low-level fct utilities.
/// </summary>
public class ExcelUtils
{
    // 1,048,576
    public const int ExcelRowNumMax = 1048576;

    /// <summary>
    /// Split a list of excel columns or range of columns.
    /// Check the validity.
    /// exp: A;B-D,G;H
    /// </summary>
    /// <param name="str"></param>
    /// <param name="listRangeCols"></param>
    /// <returns></returns>
    public static bool DecodeListRangeCols(string str, out ExcelListColsOrRangeCols listRangeCols)
    {
        listRangeCols = new ExcelListColsOrRangeCols();

        if (string.IsNullOrEmpty(str))
            // empty, not an error
            return true;

        string[] arrColRange = str.Split(';');
        // scan each item
        foreach (string s in arrColRange)
        {
            string item = s.Trim();

            // is it a col or a range of col?
            if (item.Contains("-"))
            {
                // decode the range
                if (!DecodeRangeCols(item, out ExcelRangeCols excelRangeCols))
                    return false;
                listRangeCols.ListRangeCols.Add(excelRangeCols);
                continue;
            }
            // its a unique column, decode it
            int colIdx = ColumnNameToNumber(item);
            if (colIdx < 0) return false;

            // save
            ExcelCol col = new ExcelCol(item, colIdx);
            listRangeCols.ListRangeCols.Add(col);
        }

        // check the content, columns should ordered 
        if (!CheckExcelListRangeColsIsOrdered(listRangeCols))
            return false;

        return true;
    }

    public static bool DecodeRangeCols(string item, out ExcelRangeCols excelRangeCols)
    {
        excelRangeCols = null;
        string[] arrItems = item.Split("-");

        // should have 3 parts
        if (arrItems.Count() != 2)
            return false;

        // decode the col start
        int colStartIdx = ColumnNameToNumber(arrItems[0]);
        if (colStartIdx < 0) return false;


        // decode the col end
        int colEndIdx = ColumnNameToNumber(arrItems[1]);
        if (colEndIdx < 0) return false;

        // save the range
        excelRangeCols = new ExcelRangeCols(arrItems[0], colStartIdx, arrItems[1], colEndIdx);
        return true;

    }

    /// <summary>
    /// Convert an Excel column name (letter) to a number.
    /// 
    /// Return the value of the column in base1.
    /// return -1 if an the string is wrong.
    /// return -2 if the col value is out of range.
    /// 
    /// 'A' the expected result will be 1
    /// 'AH' = 34
    /// 'XFD' = 16384
    /// </summary>
    /// <param name="columnName"></param>
    /// <returns></returns>
    public static int ColumnNameToNumber(string columnName)
    {
        if (string.IsNullOrEmpty(columnName))
            // error
            return -1;

        columnName = columnName.ToUpperInvariant();

        int sum = 0;

        for (int i = 0; i < columnName.Length; i++)
        {
            // should be a letter
            if (!char.IsLetter(columnName[i]))
                return 0;

            sum *= 26;
            sum += (columnName[i] - 'A' + 1);
        }

        if (sum > 16384)
            // out of range!
            return -2;

        return sum;
    }

    /// <summary>
    /// return the column name (letter) and the column index.
    /// 
    /// exp: 
    /// B6 -> return B,2
    /// AB3 -> return AB, 28
    /// </summary>
    /// <param name="colRowName"></param>
    /// <param name="colName"></param>
    /// <param name="colIndex"></param>
    /// <returns></returns>
    public static bool SplitCellAddress(string colRowName, out string colName, out int colIndex, out int rowIndex)
    {
        colName = "";
        colIndex = 0;
        rowIndex = 0;

        if (string.IsNullOrEmpty(colRowName)) return false;
        colRowName = colRowName.ToUpperInvariant();

        // remove the row index
        for (int i = 0; i < colRowName.Length; i++)
        {
            // not a letter, continue
            if (!char.IsLetter(colRowName[i])) break;
            colName = String.Concat(colName, colRowName[i]);
        }

        if (colName == string.Empty) return false;

        int sum = 0;

        for (int i = 0; i < colName.Length; i++)
        {
            sum *= 26;
            sum += (int)(colName[i] - 'A' + 1);
        }
        colIndex = sum;


        // get the row index
        string rowStr = colRowName.Remove(0, colName.Length);

        // check the digits part
        for (int i = colName.Length; i < colRowName.Length; i++)
        {
            // not a digit, continue
            if (!char.IsDigit(colRowName[i]))
                return false;
        }

        if (rowStr.Length == 0)
            return false;
        if (!int.TryParse(rowStr, out rowIndex))
            return false;

        return true;
    }

    public static bool CheckMaxColAndRowValue(int colIndex, int rowIndex)
    {
        // XFD
        if (colIndex > 16384) return false;
        // 1,048,576
        if (rowIndex > 1048576) return false;

        return true;
    }

    /// <summary>
    /// Convert to a standard excel address.
    /// exp: 1,1 -> A1
    /// </summary>
    /// <param name="col"></param>
    /// <param name="rowIndex"></param>
    /// <returns></returns>
    public static string ConvertAddress(int colIndex, int rowIndex)
    {
        if (colIndex < 1) return string.Empty;
        if (rowIndex < 1) return string.Empty;

        return GetColumnName(colIndex) + rowIndex.ToString();
    }

    /// <summary>
    /// return the column name of the col index.
    /// exp: 1 -> A
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    public static string GetColumnName(int index)
    {
        if (index < 1) return String.Empty;

        index--;
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        var value = "";

        if (index >= letters.Length)
            value += letters[index / letters.Length - 1];

        value += letters[index % letters.Length];

        return value;
    }

    public static bool IsExcelCellAddress(string cellAddress)
    {
        if (!ExcelUtils.SplitCellAddress(cellAddress, out string colName, out int colIndex, out int rowIndex))
            return false;

        return CheckMaxColAndRowValue(colIndex, rowIndex);
    }

    /// <summary>
    /// check the content, columns should ordered.
    /// </summary>
    /// <param name="listRangeCells"></param>
    /// <returns></returns>
    public static bool CheckExcelListRangeColsIsOrdered(ExcelListColsOrRangeCols listRangeCells)
    {
        int currCol = 0;

        foreach (var colsBase in listRangeCells.ListRangeCols)
        {

            // its a Col
            ExcelCol excelCol = colsBase as ExcelCol;
            if (excelCol != null)
            {
                if (excelCol.ColumnInt < currCol)
                    // error, should be greater
                    return false;
                currCol = excelCol.ColumnInt;
                continue;
            }

            // its a range of cols
            ExcelRangeCols excelRangeCols = colsBase as ExcelRangeCols;
            if (excelRangeCols != null)
            {
                // greater
                if (excelRangeCols.ColumnStartInt < currCol)
                    return false;
                currCol = excelRangeCols.ColumnStartInt;

                if (excelRangeCols.ColumnEndInt < currCol)
                    return false;
                currCol = excelRangeCols.ColumnEndInt;
            }
        }

        return true;
    }

}
