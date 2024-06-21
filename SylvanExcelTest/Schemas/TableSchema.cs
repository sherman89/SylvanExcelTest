using Sylvan.Data;
using Sylvan.Data.Excel;
using SylvanExcelTest.Shared;

namespace SylvanExcelTest.Schemas;

public class TableSchema(string worksheetName, Schema schema, Language language)
{
    public string WorksheetName { get; set; } = worksheetName;

    public Schema Schema { get; set; } = schema;

    public Language Language { get; set; } = language;

    public virtual BaseValidator? GetValidator(ExcelDataReader edr, List<string> errors)
    {
        return new DefaultValidator(errors);
    }
}