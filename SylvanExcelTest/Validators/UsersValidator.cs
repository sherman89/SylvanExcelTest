using System.Data.Common;
using Sylvan.Data;
using Sylvan.Data.Excel;
using SylvanExcelTest.Shared;

namespace SylvanExcelTest.Validators;

internal class UsersValidator : BaseValidator
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