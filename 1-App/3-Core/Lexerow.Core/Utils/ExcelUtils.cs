using Lexerow.Core.System;

namespace Lexerow.Core.Utils;

/// <summary>
/// Excel low-level fct utilities.
/// </summary>
public class ExcelUtils
{
    // 1,048,576
    public const int ExcelRowNumMax = 1048576;

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
}