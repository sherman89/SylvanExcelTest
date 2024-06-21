using Sylvan.Data.Excel;

namespace SylvanExcelTest.Shared;

/// <summary>
/// Convenient data structure for holding value and metadata of a cell.
/// Primarily used for code reusability in error reporting.
/// <br/>
/// <b>NOTE: It is assumed that the reader is at the correct row number.</b>
/// </summary>
public readonly record struct CellInfo
{
    public CellInfo(ExcelDataReader edr, int colOrdinal)
    {
        var colSchema = edr.GetColumnSchema();

        // BaseColumnName should contain the name as it appears in the Excel sheet
        // as opposed to ColumnName which holds the internal name for mapping to records
        var colName = colSchema[colOrdinal].BaseColumnName;
        if (string.IsNullOrWhiteSpace(colName))
        {
            throw new ArgumentException("Column name cannot be empty!", nameof(edr));
        }

        if (string.IsNullOrWhiteSpace(edr.WorksheetName))
        {
            throw new ArgumentException("Worksheet name cannot be empty!", nameof(edr));
        }

        Value = edr.GetString(colOrdinal); // GetString is always safe regardless of type
        WorksheetName = edr.WorksheetName;
        ColumnName = colName;
        Ordinal = colOrdinal;
        RowNumber = edr.RowNumber;
        ExcelColumnPosition = $"{ExcelHelpers.GetExcelColumnName(Ordinal + 1)}:{RowNumber}";
    }

    public string Value { get; }

    public string WorksheetName { get; }

    public string ColumnName { get; }

    public int Ordinal { get; }

    public int RowNumber { get; }

    public string ExcelColumnPosition { get; }
}