using Sylvan.Data;
using Sylvan.Data.Excel;
using SylvanExcelTest.Shared;
using SylvanExcelTest.Validators;

namespace SylvanExcelTest.Schemas;

public class EmailSchema(string worksheetName, Schema schema, Language language)
    : TableSchema(worksheetName, schema, language)
{
    public override BaseValidator? GetValidator(ExcelDataReader edr, List<string> errors)
    {
        return EmailValidator.TryCreate(edr, Schema, errors);
    }
}