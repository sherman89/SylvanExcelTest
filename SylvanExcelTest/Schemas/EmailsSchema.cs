using Sylvan.Data;
using Sylvan.Data.Excel;
using SylvanExcelTest.Shared;
using System.Data.Common;

namespace SylvanExcelTest.Schemas;

public record EmailRecord
{
    public int Id { get; set; }
    public int UserId { get; set; }
    public string? Email { get; set; }
}

internal class EmailsSchema : BaseSchemaWrapper
{
    public const string WorksheetEn = "E-mails";
    public const string WorksheetFi = "Sähköpostit";

    // Just to test worksheet existence validation :)
    private const string IAMERROR = "I AM ERROR";

    public EmailsSchema()
    {
        WorksheetSchemas.Add(WorksheetEn, new Schema.Builder()
            .Add("Id", typeof(int))
            .Add("UserId", nameof(EmailRecord.UserId), typeof(string))
            .Add("E-mail address", nameof(EmailRecord.Email), typeof(string))
            .Build());

        WorksheetNamesByLanguage.Add(WorksheetEn, Language.English);

        WorksheetSchemas.Add(WorksheetFi, new Schema.Builder()
            .Add("Id", typeof(int))
            .Add("KäyttäjäId", nameof(EmailRecord.UserId), typeof(string))
            .Add("Sähköposti", nameof(EmailRecord.Email), typeof(string))
            .Build());

        WorksheetNamesByLanguage.Add(WorksheetFi, Language.Finnish);

        WorksheetSchemas.Add(IAMERROR, new Schema.Builder()
            .Add("Id", typeof(int))
            .Add("UserId", nameof(EmailRecord.UserId), typeof(string))
            .Add("E-mail address", nameof(EmailRecord.Email), typeof(string))
            .Build());

        WorksheetNamesByLanguage.Add(IAMERROR, Language.English);
    }

    public override BaseValidator? GetValidator(ExcelDataReader edr, List<string> errors)
    {
        if (edr.WorksheetName == null)
        {
            throw new ArgumentException($"{nameof(edr.WorksheetName)} cannot be null!");
        }

        return WorksheetSchemas.TryGetValue(edr.WorksheetName, out var schema)
            ? EmailValidator.TryCreate(edr, schema, errors)
            : null;
    }
}

public class EmailValidator : BaseValidator
{
    private readonly int _idOrd, _userOrd, _emailOrd;

    public static EmailValidator? TryCreate(
        DbDataReader reader,
        Schema schema,
        List<string> errors)
    {
        return ValidateColumnNames(reader, schema, errors)
            ? new EmailValidator(reader, schema, errors)
            : null;
    }

    private EmailValidator(DbDataReader reader, Schema schema, List<string> errors) : base(errors)
    {
        _idOrd = reader.GetOrdinal(schema[0].ColumnName);
        _userOrd = reader.GetOrdinal(schema[1].ColumnName);
        _emailOrd = reader.GetOrdinal(schema[2].ColumnName);
    }

    protected override bool ValidateCustom(DataValidationContext context)
    {
        var edr = (ExcelDataReader)context.DataReader;
        var isValid = true;

        var email = edr.GetString(_emailOrd);
        if (!email.Contains('@'))
        {
            LogError(context, _emailOrd);
            isValid = false;
        }

        return isValid;
    }
}