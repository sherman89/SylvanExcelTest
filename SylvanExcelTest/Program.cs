using Sylvan.Data;
using Sylvan.Data.Excel;
using System.Data.Common;
using SylvanExcelTest;

internal class Program
{
    private static int Main()
    {
        // TODO: Exception handling (report unbound columns/properties)
        // TODO: Handle GetOrdinal errors in validators

        const string userNameEn = "Users";
        var userSchemaEn = new Schema.Builder()
            .Add("Id", typeof(int))
            .Add("First name", nameof(UserRecord.FirstName), typeof(string))
            .Add("Last name", nameof(UserRecord.LastName), typeof(string))
            .Add("Age", typeof(int))
            .Build();

        const string userNameFi = "Käyttäjät";
        var userSchemaFi = new Schema.Builder()
            .Add("Id", typeof(int))
            .Add("Etunimi", nameof(UserRecord.FirstName), typeof(string))
            .Add("Sukunimi", nameof(UserRecord.LastName), typeof(string))
            .Add("Ikä", typeof(int))
            .Build();

        const string emailNameEn = "E-mails";
        var emailSchemaEn = new Schema.Builder()
            .Add("Id", typeof(int))
            .Add("UserId", nameof(EmailRecord.UserId), typeof(string))
            .Add("E-mail address", nameof(EmailRecord.Email), typeof(string))
            .Build();

        const string emailNameFi = "Sähköpostit";
        var emailSchemaFi = new Schema.Builder()
            .Add("Id", typeof(int))
            .Add("KäyttäjäId", nameof(EmailRecord.UserId), typeof(string))
            .Add("Sähköposti", nameof(EmailRecord.Email), typeof(string))
            .Build();


        var schema = new ExcelSchema();
        schema.Add(userNameEn, true, userSchemaEn);
        schema.Add(emailNameEn, true, emailSchemaEn);
        schema.Add(userNameFi, true, userSchemaFi);
        schema.Add(emailNameFi, true, emailSchemaFi);

        var opts = new ExcelDataReaderOptions { Schema = schema };

        // this new file has some extra validation examples
        using var edr = ExcelDataReader.Create("TestData/Data_ENG2.xlsx", opts);

        // TODO: Validate that all expected worksheets are found?
        // var worksheetNames = edr.WorksheetNames;

        // These collections will receive the valid data.
        List<UserRecord>? users = null;
        List<EmailRecord>? emails = null;

        var errors = new List<string>();

        // TODO: Possible to refactor and simplify below even more?

        do
        {
            Validator validator;
            switch (edr.WorksheetName)
            {
                case userNameEn:
                    if (!Validator.ValidateSchema(edr, userSchemaEn, errors))
                    {
                        continue;
                    }
                    validator = new UserValidator(edr, userSchemaEn);
                    break;
                case userNameFi:
                    if (!Validator.ValidateSchema(edr, userSchemaFi, errors))
                    {
                        continue;
                    }
                    validator = new UserValidator(edr, userSchemaFi);
                    break;
                case emailNameEn:
                    if (!Validator.ValidateSchema(edr, emailSchemaEn, errors))
                    {
                        continue;
                    }
                    validator = new EmailValidator(edr, emailSchemaEn);
                    break;
                case emailNameFi:
                    if (!Validator.ValidateSchema(edr, emailSchemaFi, errors))
                    {
                        continue;
                    }
                    validator = new EmailValidator(edr, emailSchemaFi);
                    break;
                default:
                    continue;
            }

            var reader = edr.Validate((context) => validator.Validate(context, errors));

            try
            {
                switch (validator)
                {
                    case UserValidator:
                        users = reader.GetRecords<UserRecord>().ToList();
                        break;
                    case EmailValidator:
                        emails = reader.GetRecords<EmailRecord>().ToList();
                        break;
                }
            }
            catch (UnboundMemberException ume)
            {
                // AFAIK In this setup it's only possible to get this if there's a bug in the code
                Console.WriteLine(ume);
            }
        } while (edr.NextResult());

        // output errors
        Console.WriteLine("Errors:");
        errors.ForEach(Console.WriteLine);

        // output the valid records.
        Console.WriteLine();
        Console.WriteLine("Valid results:");
        if (users != null)
        {
            Console.WriteLine("Users:");
            foreach (var user in users)
            {
                Console.WriteLine(user);
            }
        }

        if (emails != null)
        {
            Console.WriteLine("Emails:");
            foreach (var email in emails)
            {
                Console.WriteLine(email);
            }
        }

        return 0;
    }

    private abstract class Validator
    {
        public static bool ValidateSchema(DbDataReader reader, Schema schema, List<string> errors)
        {
            var excelSchema = reader.GetColumnSchema();

            foreach (var dbColumn in schema)
            {
                // If current columns do not contain expected column name, log error
                if (excelSchema.All(c => c.BaseColumnName != dbColumn.BaseColumnName))
                {
                    errors.Add($"Expected to find column \"{dbColumn.BaseColumnName}\" but did not.");
                    return false;
                }
            }

            return true;
        }

        public bool Validate(DataValidationContext context, List<string> errors)
        {
            var edr = (ExcelDataReader)context.DataReader;

            var valid = LogSchemaErrors(context, edr, errors);
            valid &= ValidateCustom(context, edr, errors);

            return valid;
        }

        protected abstract bool ValidateCustom(DataValidationContext context, ExcelDataReader edr, List<string> errors);

        protected void LogError(ExcelDataReader dr, int ord, List<string> errors)
        {
            var name = dr.GetName(ord);
            var value = dr.GetString(ord); // always safe to GetString
            var type = dr.GetDataTypeName(ord);

            var colPosition = ExcelHelpers.GetExcelColumnName(ord + 1);
            errors.Add($"Invalid {name} at position \"{colPosition}:{dr.RowNumber}\" value '{value}'.");
        }

        private bool LogSchemaErrors(DataValidationContext context, ExcelDataReader dr, List<string> errors)
        {
            var valid = true;
            foreach (var ord in context.GetErrors())
            {
                LogError(dr, ord, errors);
                valid = false;
            }
            return valid;
        }
    }

    private class UserValidator : Validator
    {
        private readonly int _idOrd;
        private readonly int _fnOrd;
        private readonly int _lnOrd;
        private readonly int _ageOrd;

        private readonly HashSet<int> _seenIds = [];

        public UserValidator(DbDataReader reader, Schema schema)
        {
            _idOrd = reader.GetOrdinal(schema[0].ColumnName);
            _fnOrd = reader.GetOrdinal(schema[1].ColumnName);
            _lnOrd = reader.GetOrdinal(schema[2].ColumnName);
            _ageOrd = reader.GetOrdinal(schema[3].ColumnName);
        }

        protected override bool ValidateCustom(DataValidationContext context, ExcelDataReader dr, List<string> errors)
        {
            var valid = true;
            
            if (context.IsValid(_idOrd))
            {
                var id = context.GetValue<int>(_idOrd);
                if (!_seenIds.Add(id))
                {
                    var colPosition = ExcelHelpers.GetExcelColumnName(_idOrd + 1);
                    errors.Add($"Duplicate Id at position \"{colPosition}:{dr.RowNumber}\" value '{id}'.");
                    valid = false;
                }
            }

            var fn = dr.GetString(_fnOrd);
            if (string.IsNullOrEmpty(fn))
            {
                LogError(dr, _fnOrd, errors);
                valid = false;
            }

            var ln = dr.GetString(_lnOrd);
            if (string.IsNullOrEmpty(ln))
            {
                LogError(dr, _lnOrd, errors);
                valid = false;
            }

            // Note: default value is returned if cell has invalid string. Potentially an issue but probably not.
            var age = context.GetValue<int>(_ageOrd);
            if (age < 0)
            {
                LogError(dr, _ageOrd, errors);
                valid = false;
            }

            return valid;
        }
    }

    private class EmailValidator : Validator
    {
        private readonly int _idOrd, _userOrd, _emailOrd;

        public EmailValidator(DbDataReader reader, Schema schema)
        {
            // TODO: Log exceptions
            _idOrd = reader.GetOrdinal(schema[0].ColumnName);
            _userOrd = reader.GetOrdinal(schema[1].ColumnName);
            _emailOrd = reader.GetOrdinal(schema[2].ColumnName);
        }

        protected override bool ValidateCustom(DataValidationContext context, ExcelDataReader edr, List<string> errors)
        {
            var isValid = true;

            var email = edr.GetString(_emailOrd);
            if (!email.Contains('@'))
            {
                LogError(edr, _emailOrd, errors);
                isValid = false;
            }

            return isValid;
        }
    }

}

internal static class ExtensionMethods
{
    public static Schema.Builder Add(this Schema.Builder builder, string name, Type type)
    {
        return builder.Add(new Schema.Column.Builder { BaseColumnName = name, ColumnName = name, DataType = type });
    }

    public static Schema.Builder Add(this Schema.Builder builder, string? baseName, string name, Type type)
    {
        return builder.Add(new Schema.Column.Builder { BaseColumnName = baseName, ColumnName = name, DataType = type });
    }
}