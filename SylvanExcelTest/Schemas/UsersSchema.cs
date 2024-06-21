using Sylvan.Data.Excel;
using Sylvan.Data;
using SylvanExcelTest.Shared;
using System.Data.Common;

namespace SylvanExcelTest.Schemas;

public record UserRecord
{
    public int Id { get; set; }
    public string? FirstName { get; set; }
    public string? LastName { get; set; }
    public int Age { get; set; }
    public string? JobTitle { get; set; }
}

public class UsersSchema : BaseSchemaWrapper
{
    public const string WorksheetEn = "Users";
    public const string WorksheetFi = "Käyttäjät";

    private readonly CodeLookupRepository _codeLookup;

    public UsersSchema(CodeLookupRepository codeLookup)
    {
        _codeLookup = codeLookup;

        WorksheetSchemas.Add(WorksheetEn, new Schema.Builder()
            .Add("Id", nameof(UserRecord.Id), typeof(int))
            .Add("First name", nameof(UserRecord.FirstName), typeof(string))
            .Add("Last name", nameof(UserRecord.LastName), typeof(string))
            .Add("Age", nameof(UserRecord.Age), typeof(int))
            .Add("Job Title", nameof(UserRecord.JobTitle), typeof(string))
            .Build());

        WorksheetSchemas.Add(WorksheetFi, new Schema.Builder()
            .Add("Id", nameof(UserRecord.Id), typeof(int))
            .Add("Etunimi", nameof(UserRecord.FirstName), typeof(string))
            .Add("Sukunimi", nameof(UserRecord.LastName), typeof(string))
            .Add("Ikä", nameof(UserRecord.Age), typeof(int))
            .Add("Työnimike", nameof(UserRecord.JobTitle), typeof(string))
            .Build());
    }

    public override BaseValidator? GetValidator(ExcelDataReader edr, List<string> errors)
    {
        if (edr.WorksheetName == null)
        {
            throw new ArgumentException($"{nameof(edr.WorksheetName)} cannot be null!");
        }

        var worksheetLanguage = edr.WorksheetName.Equals(WorksheetFi, StringComparison.OrdinalIgnoreCase)
            ? Language.Finnish
            : Language.English;

        return WorksheetSchemas.TryGetValue(edr.WorksheetName, out var schema)
            ? UsersValidator.TryCreate(edr, schema, errors, _codeLookup, worksheetLanguage)
            : null;
    }
}

public class UsersValidator : BaseValidator
    {
        private readonly int _idOrdinal, _firstNameOrdinal, _lastNameOrdinal, _ageOrdinal, _jobTitleOrdinal;

        private readonly CodeLookupRepository _codeLookup;
        private readonly Language _worksheetLanguage;

        private readonly HashSet<int> _seenIds = [];

        public static UsersValidator? TryCreate(
            DbDataReader reader,
            Schema schema,
            List<string> errors,
            CodeLookupRepository codeLookup,
            Language lang)
        {
            return ValidateColumnNames(reader, schema, errors)
                ? new UsersValidator(reader, schema, errors, codeLookup, lang)
                : null;
        }

        private UsersValidator(
            DbDataReader reader, 
            Schema schema,
            List<string> errors,
            CodeLookupRepository codeLookup,
            Language lang) : base(errors)
        {
            _codeLookup = codeLookup;
            _worksheetLanguage = lang;

            _idOrdinal = reader.GetOrdinal(schema[0].ColumnName);
            _firstNameOrdinal = reader.GetOrdinal(schema[1].ColumnName);
            _lastNameOrdinal = reader.GetOrdinal(schema[2].ColumnName);
            _ageOrdinal = reader.GetOrdinal(schema[3].ColumnName);
            _jobTitleOrdinal = reader.GetOrdinal(schema[4].ColumnName);
        }

        protected override bool ValidateCustom(DataValidationContext context)
        {
            var edr = (ExcelDataReader)context.DataReader;
            var valid = true;

            // Validate duplicate id's
            // Schema validator will invalidate the row, so we don't have to set validate to IsValid result
            if (context.IsValid(_idOrdinal))
            {
                var id = context.GetValue<int>(_idOrdinal);
                if (!_seenIds.Add(id))
                {
                    LogDuplicateCellError(context, _idOrdinal);
                    valid = false;
                }
            }

            // Validate that FirstName and LastName are together empty or non-empty
            valid &= ValidateBothOrNeitherEmpty(context, _firstNameOrdinal, _lastNameOrdinal);

            // Validate age
            if (context.IsValid(_ageOrdinal))
            {
                var age = context.GetValue<int>(_ageOrdinal);
                if (age < 0)
                {
                    LogError(context, _ageOrdinal);
                    valid = false;
                }
            }

            // Validate job title is a valid dropdown value
            var jobTitleNotEmpty = ValidateEmptyString(context, _jobTitleOrdinal);
            if (jobTitleNotEmpty)
            {
                var jobTitleDescription = edr.GetString(_jobTitleOrdinal).Trim();
                var jobTitleCode = _codeLookup.GetJobTitleCode(_worksheetLanguage, jobTitleDescription);

                if (jobTitleCode == null)
                {
                    var cellInfo = new CellInfo(edr, _jobTitleOrdinal);
                    Errors.Add($"Code for job title \"{jobTitleDescription}\" was not found in our system." +
                               $"Worksheet: \"{edr.WorksheetName}\", Column: \"{cellInfo.ColumnName}\", Cell: \"{cellInfo.ExcelColumnPosition}\".");
                    valid = false;
                }
                else
                {
                    // Replace description with code, so we don't have to do it again when returning to frontend
                    context.SetValue(_jobTitleOrdinal, jobTitleCode);
                }
            }
            else
            {
                valid = false;
            }

            return valid;
        }
    }