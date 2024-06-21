using Sylvan.Data;
using Sylvan.Data.Excel;
using SylvanExcelTest.Schemas;

namespace SylvanExcelTest;

internal class Program
{
    private static async Task<int> Main()
    {
        await ParseExcelAsync("TestData/Data_ENG.xlsx");
        await ParseExcelAsync("TestData/Data_ENG2.xlsx");
        await ParseExcelAsync("TestData/Data_FIN.xlsx");

        return 0;
    }

    private static async Task ParseExcelAsync(string filePath, CancellationToken cancellationToken = default)
    {
        // NOTE: In ASP.NET, we get a stream. use filename to get workbook type and pass cancellation token.

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