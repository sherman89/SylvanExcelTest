using Sylvan.Data;
using Sylvan.Data.Excel;
using SylvanExcelTest;
using System.Data.Common;

internal class Program
{
    private static int Main()
    {
        // TODO: Exception handling (report unbound columns/properties)
        // TODO: Handle GetOrdinal errors in validators

        var userTable =
            new Table<UserValidator>("Users", ["Users", "Käyttäjät"])
                .Add<int>("Id", ["Id"], t => t.ValidateId)
                .Add<string>("FirstName", ["First name", "Etunimi"], t => t.ValidateName)
                .Add<string>("LastName", ["Last name", "Sukunimi"], t => t.ValidateName)
                .Add<int>("Age", ["Age", "Ikä"], t => t.ValidateAge);

        var emailTable =
            new Table<EmailValidator>("Emails", ["E-mails", "Sähköpostit"])
                .Add<int>("Id", ["Id"])
                .Add<int>("UserId", ["UserId", "KäyttäjäId"])
                .Add<string>("Email", ["E-mail address", "Sähköposti"], t => t.ValidateEmail);

        var schema = new ValidatingSchema() { userTable, emailTable };

        var opts = new ExcelDataReaderOptions { Schema = schema };

        // this new file has some extra validation examples
        using var edr = ExcelDataReader.Create("TestData/Data_FIN.xlsx", opts);

        // TODO: Validate that all expected worksheets are found?
        // var worksheetNames = edr.WorksheetNames;        

        // These collections will receive the valid data.
        List<UserRecord>? users = null;
        List<EmailRecord>? emails = null;

        var errors = new List<string>();

        do
        {
            var table = schema.FindTable(edr.WorksheetName);
            if (table == null)
            {
                errors.Add($"Unknown table {edr.WorksheetName}");
                continue;
            }

            ValidateSchema(table, edr, errors);

            // ask the schema for a validator for this sheet
            var validator = table.GetValidator(edr, errors);
            // apply the validator
            var validatingReader = edr.Validate(validator.Validate);

            // bind based on which table we found
            if (table == userTable)
            {
                users = validatingReader.GetRecords<UserRecord>().ToList();
            }

            if (table == emailTable)
            {
                emails = validatingReader.GetRecords<EmailRecord>().ToList();
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

    static bool ValidateSchema(Table schema, DbDataReader reader, List<string> errors)
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

    abstract class LoggingValidator : TableValidator
    {
        protected List<string> errors;

        public override void SetState(object? o)
        {
            ArgumentNullException.ThrowIfNull(o);
            this.errors = (List<string>)o;
        }

        public override void LogError(DataValidationContext context, int ord, Exception ex = null)
        {
            var dr = context.DataReader;
            var name = dr.GetName(ord);
            var value = dr.GetString(ord); // always safe to GetString
            var type = dr.GetDataTypeName(ord);
            var row = context.RowNumber;
            var colPosition = ExcelHelpers.GetExcelColumnName(ord + 1);
            errors.Add($"Invalid {name} at position {row},{ord} ({colPosition}) value '{value}'.");
        }                
    }

    class UserValidator : LoggingValidator
    {
        private readonly HashSet<int> _seenIds = [];

        // used for both first and last names
        public bool ValidateName(DataValidationContext context, int ordinal)
        {
            var name = context.DataReader.GetString(ordinal);
            return !string.IsNullOrWhiteSpace(name);
        }

        public bool ValidateId(DataValidationContext context, int ordinal)
        {
            var id = context.DataReader.GetInt32(ordinal);
            if (!_seenIds.Add(id))
            {
                var colPosition = ExcelHelpers.GetExcelColumnName(ordinal + 1);
                errors.Add($"Duplicate Id at position \"{colPosition}:{context.RowNumber}\" value '{id}'.");
                return false;
            }
            return true;
        }

        public bool ValidateAge(DataValidationContext context, int ordinal)
        {
            var age = context.DataReader.GetInt32(ordinal);
            return age > 0 && age < 120;
        }
    }

    class EmailValidator : LoggingValidator
    {
        public bool ValidateEmail(DataValidationContext context, int ordinal)
        {
            var email = context.DataReader.GetString(ordinal);
            return !email.Contains('@');
        }
    }
}
