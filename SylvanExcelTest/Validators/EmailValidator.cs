using System.Data.Common;
using Sylvan.Data;
using Sylvan.Data.Excel;
using SylvanExcelTest.Shared;

namespace SylvanExcelTest.Validators;

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