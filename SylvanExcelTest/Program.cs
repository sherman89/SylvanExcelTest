using Sylvan.Data;
using Sylvan.Data.Excel;
using System.Collections.Immutable;
using System.Collections.ObjectModel;
using System.Data.Common;
using System.Diagnostics;

namespace SylvanExcelTest
{
    internal class Program
    {
        static async Task<int> Main(string[] args)
        {
            // TODO: Maybe would be nice to use this, but how to support multiple worksheets? Have to use 2 readers
            // var userWorksheetSchema = Schema.Parse("Id:string,FirstName:string,LastName:string,Age:string");
            // var emailWorksheetSchema = Schema.Parse("Id:string,UserId:string,Email:string");
            // var excelSchema = new ExcelSchema(hasHeaders: true, userWorksheetSchema);

            var schema = new MyExcelSchema();

            try
            {
                // Get directory path to exe so we can read test xlsx files that are copied to output dir
                var procDirPath = Path.GetDirectoryName(Environment.ProcessPath);

                await using var edr = await ExcelDataReader.CreateAsync(
                    @$"{procDirPath}\TestData\Data_FIN.xlsx",
                    new ExcelDataReaderOptions
                    {
                        Schema = schema
                    });

                
                // TODO: How to get this called while getting records / reading?
                // TODO: Why edr works but not this when in GetRecordsAsync?
                var reader = edr.Validate(context =>
                {
                    // I don't know what I'm doing here :)
                    var errors = context.GetErrors().ToList();
                    Console.WriteLine($"Ordinals with errors: {string.Join(", ", errors)}");
                    return errors.Count == 0;
                });

                var result = new ResultRecord();

                do
                {
                    if (edr.WorksheetName == null)
                    {
                        continue;
                    }

                    var sheetName = schema.MapToInternalWorksheetName(edr.WorksheetName);

                    switch (sheetName)
                    {
                        case MyExcelSchema.UsersWorksheet:
                        {
                            var rowNumber = edr.RowNumber;
                            var columns = edr.GetColumnSchema();
                            result.Users = await edr.GetRecordsAsync<UserRecord>().ToArrayAsync();
                            ValidateUsers(result.Users, columns, rowNumber);
                            break;
                        }
                        case MyExcelSchema.EmailsWorksheet:
                            result.Emails = await edr.GetRecordsAsync<EmailRecord>().ToArrayAsync();
                            break;
                    }
                } while (await edr.NextResultAsync());
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return -1;
            }

            return 0;
        }

        private static void ValidateUsers(UserRecord[] data, ReadOnlyCollection<DbColumn> columns, int firstRowNumber)
        {
            // TODO: A lot of logic just to validate one column... doing this for all would not look pretty? how to simplify?

            var emptyFirstNames = data
                .Select((usr, index) => new { User = usr, index })
                .Where(grp => string.IsNullOrWhiteSpace(grp.User.FirstName))
                .ToList();

            foreach (var user in emptyFirstNames)
            {
                var firstNameCol = columns.First(c => c.ColumnName == nameof(UserRecord.FirstName));

                Debug.Assert(firstNameCol.ColumnOrdinal != null, "Schema always sets Ordinal, shouldn't be null");
                var columnPosition = ExcelHelpers.GetExcelColumnName(firstNameCol.ColumnOrdinal.Value + 1);

                Console.WriteLine($"Invalid value at position {columnPosition}:{firstRowNumber + user.index + 1}. " +
                                  $"Column: {firstNameCol.BaseColumnName}, Worksheet: {firstNameCol.BaseTableName}");
            }
        }
    }

    public record ResultRecord
    {
        public UserRecord[]? Users { get; set; }

        public EmailRecord[]? Emails { get; set; }
    }

    public record UserRecord
    {
        public string? Id { get; set; }
        public string? FirstName { get; set; }
        public string? LastName { get; set; }
        public string? Age { get; set; }
    }

    public record EmailRecord
    {
        public string? Id { get; set; }
        public string? UserId { get; set; }
        public string? Email { get; set; }
    }

    public sealed class MyExcelSchema : ExcelSchemaProvider
    {
        // _columnMapping: I could also map by name e.g. "{UsersWorksheet}.Id" but maybe better by ordinal?

        private readonly ImmutableDictionary<string, ExcelColumn> _columnMapping =
            ImmutableDictionary.CreateRange(StringComparer.OrdinalIgnoreCase,
                new KeyValuePair<string, ExcelColumn>[]
                {
                    // Users
                    new($"{UsersWorksheet}.0", UserIdCol),
                    new($"{UsersWorksheet}.1", UserFirstNameCol),
                    new($"{UsersWorksheet}.2", UserLastNameCol),
                    new($"{UsersWorksheet}.3", UserAgeCol),
                    // Emails
                    new($"{EmailsWorksheet}.0", EmailIdCol),
                    new($"{EmailsWorksheet}.1", EmailUserIdCol),
                    new($"{EmailsWorksheet}.2", EmailCol)
                });

        private readonly ImmutableDictionary<string, string> _worksheetMapping =
            ImmutableDictionary.CreateRange(StringComparer.OrdinalIgnoreCase,
                new KeyValuePair<string, string>[]
                {
                    new("Users", UsersWorksheet),
                    new("Käyttäjät", UsersWorksheet),
                    new("E-mails", EmailsWorksheet),
                    new("Sähköpostit", EmailsWorksheet)
                });

        public const string UsersWorksheet = "Users";
        private static readonly ExcelColumn UserIdCol = new (nameof(UserRecord.Id), typeof(string), 0, UsersWorksheet);
        private static readonly ExcelColumn UserFirstNameCol = new (nameof(UserRecord.FirstName), typeof(string), 1, UsersWorksheet);
        private static readonly ExcelColumn UserLastNameCol = new (nameof(UserRecord.LastName), typeof(string), 2, UsersWorksheet);
        private static readonly ExcelColumn UserAgeCol = new (nameof(UserRecord.Age), typeof(string), 3, UsersWorksheet);

        public const string EmailsWorksheet = "E-mail";
        private static readonly ExcelColumn EmailIdCol = new (nameof(EmailRecord.Id), typeof(string), 0, EmailsWorksheet);
        private static readonly ExcelColumn EmailUserIdCol = new (nameof(EmailRecord.UserId), typeof(string), 1, EmailsWorksheet);
        private static readonly ExcelColumn EmailCol = new (nameof(EmailRecord.Email), typeof(string), 2, EmailsWorksheet);

        public override DbColumn? GetColumn(string sheetName, string? name, int ordinal)
        {
            var mappedWorksheetName = MapToInternalWorksheetName(sheetName);
            if (mappedWorksheetName == null)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                var colPosition = ExcelHelpers.GetExcelColumnName(ordinal + 1);
                // TODO: Can I get row id here somehow in-case header is not first row?
                throw new ExcelReadException($"Column in worksheet \"{sheetName}\" at position \"{colPosition}\" has an empty header.", ordinal)
                {
                    WorksheetName = sheetName,
                    ColumnName = name
                };
            }

            return CollectionExtensions.GetValueOrDefault(_columnMapping, $"{mappedWorksheetName}.{ordinal}");
        }

        public override bool HasHeaders(string sheetName)
        {
            return true;
        }

        public string? MapToInternalWorksheetName(string sheetName)
        {
            return CollectionExtensions.GetValueOrDefault(_worksheetMapping, sheetName.Trim());
        }
    }

    public class ExcelColumn : DbColumn
    {
        public ExcelColumn(string name, Type type, int ordinal, string tableName, bool allowNull = false)
        {
            ColumnName = name;
            DataType = type;
            AllowDBNull = allowNull;
            ColumnOrdinal = ordinal;
            BaseTableName = tableName;
        }
    }

    public class ExcelReadException(string? message, int ordinal) : Exception(message)
    {
        public string? WorksheetName { get; set; }
        public string? ColumnName { get; set; }
        public int Ordinal { get; set; } = ordinal;
    }
}
