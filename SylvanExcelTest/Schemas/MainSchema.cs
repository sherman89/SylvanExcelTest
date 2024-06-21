using Sylvan.Data.Excel;
using SylvanExcelTest.Shared;

namespace SylvanExcelTest.Schemas;

public record MainRecord
{
    public List<UserRecord>? Users { get; set; }
    public List<EmailRecord>? Emails { get; set; }
}

public class MainSchema
{
    public readonly Dictionary<string, BaseSchemaWrapper>
        WorksheetSchemaWrappers = new(StringComparer.OrdinalIgnoreCase);

    public readonly ExcelSchema ExcelSchema = new();

    public MainSchema(CodeLookupRepository codeLookup)
    {
        var usersSchema = new UsersSchema(codeLookup);
        foreach (var kvp in usersSchema.WorksheetSchemas)
        {
            WorksheetSchemaWrappers.Add(kvp.Key, usersSchema);
            ExcelSchema.Add(kvp.Key, true, kvp.Value);
        }

        var emailsSchema = new EmailsSchema();
        foreach (var kvp in emailsSchema.WorksheetSchemas)
        {
            WorksheetSchemaWrappers.Add(kvp.Key, emailsSchema);
            ExcelSchema.Add(kvp.Key, true, kvp.Value);
        }
    }
}