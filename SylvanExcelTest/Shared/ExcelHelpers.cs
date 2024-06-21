namespace SylvanExcelTest.Shared;

public static class ExcelHelpers
{
    /// <summary>
    /// Bijective base-26 representation of given number.
    /// </summary>
    /// <param name="columnNumber">Position of column, starting from 1</param>
    public static string GetExcelColumnName(int columnNumber)
    {
        var columnName = "";

        while (columnNumber > 0)
        {
            var modulo = (columnNumber - 1) % 26;
            columnName = Convert.ToChar('A' + modulo) + columnName;
            columnNumber = (columnNumber - modulo) / 26;
        }

        return columnName;
    }
}