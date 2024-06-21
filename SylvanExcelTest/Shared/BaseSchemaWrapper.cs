using Sylvan.Data;
using Sylvan.Data.Excel;

namespace SylvanExcelTest.Shared;

public abstract class BaseSchemaWrapper
{
    public readonly Dictionary<string, Schema> WorksheetSchemas = new(StringComparer.OrdinalIgnoreCase);

    // TODO: Base class shouldn't have any idea about language I think... I don't like this...

    public readonly Dictionary<string, Language> WorksheetNamesByLanguage = new(StringComparer.OrdinalIgnoreCase);

    public abstract BaseValidator? GetValidator(ExcelDataReader edr, List<string> errors);
}