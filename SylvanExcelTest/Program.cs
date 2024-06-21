using Sylvan.Data;
using Sylvan.Data.Excel;
using SylvanExcelTest.Schemas;
using SylvanExcelTest.Shared;

namespace SylvanExcelTest;

internal class Program
{
    // Yes I know too many dictionaries and complexity... Need to refactor and keep features but brain has no better ideas right now :|

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

        var mainSchema = new MainSchema(codeLookupRepo);

        var opts = new ExcelDataReaderOptions
        {
            // Culture = new CultureInfo("fi-FI"),
            OwnsStream = true,
            Schema = mainSchema.ExcelSchema
        };

        await using var edr = await ExcelDataReader.CreateAsync(filePath, opts, cancellationToken);

        var result = new MainRecord();
        var errors = new List<string>();

        // Ensure that worksheets we expect are found
        ValidateWorksheet(
            mainSchema.GetWorksheetNamesByLanguage(language), 
            edr.WorksheetNames.ToList(), 
            errors);

        do
        {
            if (edr.WorksheetName == null)
            {
                continue;
            }

            mainSchema.WorksheetSchemaWrappers.TryGetValue(edr.WorksheetName, out var schemaWrapper);

            var validator = schemaWrapper?.GetValidator(edr, errors);
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

    private static void ValidateWorksheet(List<string> expectedWorksheets, List<string> actualWorksheets, List<string> errors)
    {
        foreach (var expectedWorksheet in expectedWorksheets)
        {
            if (!actualWorksheets.Contains(expectedWorksheet, StringComparer.OrdinalIgnoreCase))
            {
                errors.Add($"Expected to find worksheet \"{expectedWorksheet}\" but did not.");
            }
        }
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