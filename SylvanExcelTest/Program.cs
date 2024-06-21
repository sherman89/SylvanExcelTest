using System.Globalization;
using Sylvan.Data;
using Sylvan.Data.Excel;
using SylvanExcelTest.Records;
using SylvanExcelTest.Schemas;
using SylvanExcelTest.Shared;
using SylvanExcelTest.Validators;

namespace SylvanExcelTest;

internal class Program
{
    private static async Task<int> Main()
    {
        // Language is the language the user has selected in UI

        // The only practical reason I "need" it is to know which worksheet to expect
        // so if it's missing I can tell the user about it. If I don't give the language
        // then I'd either have to say "Worksheet FIN / ENG were missing" or give a generic
        // error that tells nothing. Probably this is the best way to deal with this...?

        // Aside from the above, once the worksheet is found, the language obviously can be
        // determined from that so the parameter is no longer needed. Maybe it can be used
        // to clean up the code some more or something though...

        await ParseExcelAsync("TestData/Data_ENG.xlsx", Language.English);
        await ParseExcelAsync("TestData/Data_ENG2.xlsx", Language.English);
        await ParseExcelAsync("TestData/Data_FIN.xlsx", Language.Finnish);

        return 0;
    }

    private static async Task ParseExcelAsync(string filePath, Language language, CancellationToken cancellationToken = default)
    {
        // NOTE: In ASP.NET, we get a stream to worm with so pass filename to get workbook type for the reader.

        var codeLookupRepo = new CodeLookupRepository();

        var usersSchemaEn = new UserSchema("Users",
            new Schema.Builder()
                .Add("Id", nameof(UserRecord.Id), typeof(int))
                .Add("First name", nameof(UserRecord.FirstName), typeof(string))
                .Add("Last name", nameof(UserRecord.LastName), typeof(string))
                .Add("Age", nameof(UserRecord.Age), typeof(int))
                .Add("Job Title", nameof(UserRecord.JobTitle), typeof(string))
                .Build(),
            Language.English,
            codeLookupRepo);

        var usersSchemaFi = new UserSchema("Käyttäjät",
            new Schema.Builder()
                .Add("Id", nameof(UserRecord.Id), typeof(int))
                .Add("Etunimi", nameof(UserRecord.FirstName), typeof(string))
                .Add("Sukunimi", nameof(UserRecord.LastName), typeof(string))
                .Add("Ikä", nameof(UserRecord.Age), typeof(int))
                .Add("Työnimike", nameof(UserRecord.JobTitle), typeof(string))
                .Build(),
            Language.Finnish,
            codeLookupRepo);

        var emailSchemaEn = new EmailSchema("E-mails",
            new Schema.Builder()
                .Add("Id", typeof(int))
                .Add("UserId", nameof(EmailRecord.UserId), typeof(string))
                .Add("E-mail address", nameof(EmailRecord.Email), typeof(string))
                .Build(),
            Language.English);

        var emailSchemaFi = new EmailSchema("Sähköpostit",
            new Schema.Builder()
                .Add("Id", typeof(int))
                .Add("KäyttäjäId", nameof(EmailRecord.UserId), typeof(string))
                .Add("Sähköposti", nameof(EmailRecord.Email), typeof(string))
                .Build(),
            Language.Finnish);

        var errorEn = new TableSchema("ERROR (EN)", Schema.Parse(""), Language.English);
        var errorFi = new TableSchema("ERROR (FI)", Schema.Parse(""), Language.Finnish);

        var schemas = new Dictionary<string, TableSchema>(StringComparer.OrdinalIgnoreCase)
        {
            { usersSchemaEn.WorksheetName, usersSchemaEn },
            { usersSchemaFi.WorksheetName, usersSchemaFi },
            { emailSchemaEn.WorksheetName, emailSchemaEn },
            { emailSchemaFi.WorksheetName, emailSchemaFi },
            { errorEn.WorksheetName, errorEn },
            { errorFi.WorksheetName, errorFi }
        };

        var excelSchema = new ExcelSchema();
        foreach (var tableSchema in schemas)
        {
            excelSchema.Add(tableSchema.Key, true, tableSchema.Value.Schema);
        }

        var opts = new ExcelDataReaderOptions
        {
            Culture = new CultureInfo("fi-FI"),
            OwnsStream = true,
            Schema = excelSchema
        };

        await using var edr = await ExcelDataReader.CreateAsync(filePath, opts, cancellationToken);

        var result = new MainRecord();
        var errors = new List<string>();

        // Ensure that worksheets we expect are found. A bit overkill yes, don't write me.
        // We know which schemas are "the same" except for language, so I pair them up
        // so that when validating, if for example the UI language is English then I should
        // expect "Users", but instead I get "Käyttäjät", then that's OK.
        // Also, I can tell users which one I expect if neither "Users" nor "Käyttäjät"
        // are found. So if English: "Expected to find worksheet Users but did not"

        // Group schemas together by language to enable robust worksheet validation.
        var similarSchemasPairedByLanguage = new List<Tuple<string, string>>
        {
            new(usersSchemaEn.WorksheetName, usersSchemaFi.WorksheetName),
            new(emailSchemaEn.WorksheetName, emailSchemaFi.WorksheetName),
            new(errorEn.WorksheetName, errorFi.WorksheetName)
        };

        var docWorksheets = edr.WorksheetNames.ToList();
        foreach (var t in similarSchemasPairedByLanguage)
        {
            if (docWorksheets.Contains(t.Item1, StringComparer.OrdinalIgnoreCase) ||
                docWorksheets.Contains(t.Item2, StringComparer.OrdinalIgnoreCase))
            {
                continue;
            }

            var expectedWorksheet = language == Language.English ? t.Item1 : t.Item2;
            errors.Add($"Expected to find worksheet \"{expectedWorksheet}\" but did not.");
        }

        do
        {
            if (edr.WorksheetName == null)
            {
                continue;
            }

            schemas.TryGetValue(edr.WorksheetName, out var table);

            var validator = table?.GetValidator(edr, errors);
            if (validator == null)
            {
                continue;
            }

            var reader = edr.Validate(validator.Validate);

            switch (validator)
            {
                case UsersValidator:
                    result.Users = await reader.GetRecordsAsync<UserRecord>().ToListAsync(cancellationToken);
                    break;
                case EmailValidator:
                    result.Emails = await reader.GetRecordsAsync<EmailRecord>().ToListAsync(cancellationToken);
                    break;
                default:
                    continue;
            }


        } while (await edr.NextResultAsync(cancellationToken));

        await edr.CloseAsync();

        PrintResults(result, errors, filePath);
    }

    private static void PrintResults(MainRecord result, List<string> errors, string filePath)
    {
        Console.WriteLine($"- Results for file: {Path.GetFileName(filePath)} -");
        Console.WriteLine();

        // output errors
        Console.WriteLine("Errors:");
        errors.ForEach(Console.WriteLine);

        // output the valid records.
        Console.WriteLine();
        Console.WriteLine("Valid results:");
        if (result.Users != null)
        {
            Console.WriteLine("Users:");
            foreach (var user in result.Users)
            {
                Console.WriteLine(user);
            }
        }

        if (result.Emails != null)
        {
            Console.WriteLine("Emails:");
            foreach (var email in result.Emails)
            {
                Console.WriteLine(email);
            }
        }

        Console.WriteLine();
        Console.WriteLine("---");
        Console.WriteLine();
    }
}