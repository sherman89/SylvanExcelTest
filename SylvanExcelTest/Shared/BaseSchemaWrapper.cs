using Sylvan.Data;
using Sylvan.Data.Excel;

namespace SylvanExcelTest.Shared;

public abstract class BaseSchemaWrapper
{
    public readonly Dictionary<string, Schema> WorksheetSchemas = new(StringComparer.OrdinalIgnoreCase);

    public abstract BaseValidator? GetValidator(ExcelDataReader edr, List<string> errors);
}