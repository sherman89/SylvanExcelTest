using Sylvan.Data;
using Sylvan.Data.Excel;
using SylvanExcelTest.Shared;
using SylvanExcelTest.Validators;

namespace SylvanExcelTest.Schemas;

public class UserSchema(string worksheetName, Schema schema, Language language, CodeLookupRepository codeLookup)
    : TableSchema(worksheetName, schema, language)
{
    public CodeLookupRepository CodeLookupRepository { get; set; } = codeLookup;

    public override BaseValidator? GetValidator(ExcelDataReader edr, List<string> errors)
    {
        return UsersValidator.TryCreate(edr, Schema, errors, CodeLookupRepository, Language);
    }
}