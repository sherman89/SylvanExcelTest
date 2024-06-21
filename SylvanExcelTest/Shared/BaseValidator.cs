using System.Data.Common;
using Sylvan.Data;
using Sylvan.Data.Excel;

namespace SylvanExcelTest.Shared;

public abstract class BaseValidator(List<string> errors)
{
    public List<string> Errors { get; init; } = errors;

    /// <summary>
    /// Checks if reader schema contains the columns that we expect in our schema (by name).
    /// If they're not found, errors are written to errors list.
    /// </summary>
    /// <param name="reader">Reader that can access excel file</param>
    /// <param name="schema">Schema as defined in code (what we expect)</param>
    /// <param name="errors">Not found columns reported here</param>
    /// <returns>True if all expected columns found, otherwise false.</returns>
    public static bool ValidateColumnNames(DbDataReader reader, Schema schema, List<string> errors)
    {
        var excelSchema = reader.GetColumnSchema();

        foreach (var expectedColumn in schema)
        {
            // If current columns do not contain expected column name, log error
            if (excelSchema.Any(c =>
                    string.Equals(c.BaseColumnName, expectedColumn.BaseColumnName,
                        StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            errors.Add($"Expected to find column \"{expectedColumn.BaseColumnName}\" but did not.");
            return false;
        }

        return true;
    }

    public bool Validate(DataValidationContext context)
    {
        var valid = LogSchemaErrors(context);
        valid &= ValidateCustom(context);

        return valid;
    }

    protected abstract bool ValidateCustom(DataValidationContext context);

    protected bool ValidateEmptyString(DataValidationContext context, int ord)
    {
        var edr = (ExcelDataReader)context.DataReader;

        var value = edr.GetString(ord);
        if (string.IsNullOrWhiteSpace(value))
        {
            LogEmptyCellError(context, ord);
            return false;
        }

        return true;
    }

    protected bool ValidateCellSchema(DataValidationContext context, int ord)
    {
        if (!context.IsValid(ord))
        {
            LogError(context, ord);
            return false;
        }

        return true;
    }

    /// <summary>
    /// Convenience method that checks whether 2 cells are together empty or not.
    /// If both are either null or not, proceeds to validate cells against schema (type check etc)
    /// </summary>
    protected bool ValidateBothOrNeitherEmpty(DataValidationContext context, int ord1, int ord2)
    {
        var edr = (ExcelDataReader)context.DataReader;
        var cell1Info = new CellInfo(edr, ord1);
        var cell2Info = new CellInfo(edr, ord2);

        var valid = true;

        var firstNull = edr.IsDBNull(ord1);
        var secondNull = edr.IsDBNull(ord2);

        if (!firstNull && secondNull)
        {
            Errors.Add($"\"{cell2Info.ColumnName}\" cannot be empty because \"{cell1Info.ColumnName}\" has a value.");
            valid = false;
        }

        if (firstNull && !secondNull)
        {
            Errors.Add($"\"{cell1Info.ColumnName}\" cannot be empty because \"{cell2Info.ColumnName}\" has a value.");
            valid = false;
        }

        // At this point either first cell & second cell are both null or have a value
        // So now we can check whether the value is valid (correct type etc...)
        valid &= ValidateCellSchema(context, ord1);
        valid &= ValidateCellSchema(context, ord2);

        return valid;
    }

    protected void LogError(DataValidationContext context, int ord)
    {
        var edr = (ExcelDataReader)context.DataReader;
        var cellInfo = new CellInfo(edr, ord);

        Errors.Add($"Worksheet \"{edr.WorksheetName}\", column \"{cellInfo.ColumnName}\" " +
                   $"cell \"{cellInfo.ExcelColumnPosition}\" contains invalid data. Value: \"{cellInfo.Value}\"");
    }

    protected void LogEmptyCellError(DataValidationContext context, int ord)
    {
        var edr = (ExcelDataReader)context.DataReader;
        var cellInfo = new CellInfo(edr, ord);

        Errors.Add($"Worksheet \"{edr.WorksheetName}\", column \"{cellInfo.ColumnName}\" " +
                   $"cell \"{cellInfo.ExcelColumnPosition}\" should not be empty.");
    }

    protected void LogDuplicateCellError(DataValidationContext context, int ord)
    {
        var edr = (ExcelDataReader)context.DataReader;
        var cellInfo = new CellInfo(edr, ord);

        Errors.Add($"Worksheet \"{edr.WorksheetName}\", column \"{cellInfo.ColumnName}\" " +
                   $"cell \"{cellInfo.ExcelColumnPosition}\" contains duplicate data. Value: \"{cellInfo.Value}\"");
    }

    private bool LogSchemaErrors(DataValidationContext context)
    {
        var valid = true;

        foreach (var ord in context.GetErrors())
        {
            LogError(context, ord);
            valid = false;
        }

        return valid;
    }
}