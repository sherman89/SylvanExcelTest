using Sylvan.Data;
using Sylvan.Data.Excel;
using System.Data.Common;

class Program
{
    static int Main()
    {
        const string UserNameEn = "Users";
        var userSchemaEn = Schema.Parse("Id:int,First name:string,Last name:string,Age:int");
        const string UserNameFi = "Käyttäjät";
        var userSchemaFi = Schema.Parse("Id:int,Etunimi:string,Sukunimi:string,Age:int");
        const string EmailNameEn = "E-mails";
        var emailSchemaEn = Schema.Parse("Id:int,UserId:int,E-mail address:string");
        const string EmailNameFi = "Sähköpostit";
        var emailSchemaFi = Schema.Parse("Id:int,KäyttäjäId:int,Sähköposti:string");

        var schema = new ExcelSchema();
        schema.Add(UserNameEn, true, userSchemaEn);
        schema.Add(EmailNameEn, true, emailSchemaEn);
        schema.Add(UserNameFi, true, userSchemaFi);
        schema.Add(EmailNameFi, true, emailSchemaFi);

        var opts = new ExcelDataReaderOptions { Schema = schema };

        // this new file has some extra validation examples
        using var edr = ExcelDataReader.Create("TestData/Data_ENG2.xlsx", opts);
            
        do
        {
            Validator validator = null;
            switch (edr.WorksheetName)
            {
                case UserNameEn:
                    validator = new UserValidator(edr, userSchemaEn);
                    break;
                case UserNameFi:
                    validator = new UserValidator(edr, userSchemaFi);
                    break;
                case EmailNameEn:
                    validator = new EmailValidator(edr, emailSchemaEn);
                    break;
                case EmailNameFi:
                    validator = new EmailValidator(edr, emailSchemaFi);
                    break;
                default:
                    Console.WriteLine("Unknown worksheet " + edr.WorksheetName);
                    break;
            }

            if (validator != null)
            {
                var reader = edr.Validate(validator.Validate);
                while (reader.Read())
                {
                    // only valid rows are yielded in the loop
                    Console.WriteLine($"Valid {edr.WorksheetName} {edr.RowNumber}");
                }
            }

        } while (edr.NextResult());

        return 0;
    }

    abstract class Validator
    {
        public bool Validate(DataValidationContext context)
        {
            ExcelDataReader edr = (ExcelDataReader) context.DataReader;
            var valid = LogSchemaErrors(context, edr);
            valid &= ValidateCustom(context, edr);
            return valid;
        }

        protected abstract bool ValidateCustom(DataValidationContext context, ExcelDataReader edr);

        protected void LogError(ExcelDataReader dr, int ord)
        {
            var name = dr.GetName(ord);
            var value = dr.GetString(ord); // always safe to GetString
            var type = dr.GetDataTypeName(ord);
            Console.WriteLine($"Invalid {name} at row {dr.RowNumber} col {ord} value '{value}'.");
        }

        protected bool LogSchemaErrors(DataValidationContext context, ExcelDataReader dr)
        {
            var valid = true;
            foreach (var ord in context.GetErrors())
            {
                LogError(dr, ord);
                valid = false;
            }
            return valid;
        }
    }

    class UserValidator : Validator
    {
        int idOrd, fnOrd, lnOrd, ageOrd;
        
        HashSet<int> seenIds = new HashSet<int>();

        public UserValidator(DbDataReader reader, Schema schema)
        {
            idOrd = reader.GetOrdinal(schema[0].ColumnName);
            fnOrd = reader.GetOrdinal(schema[1].ColumnName);
            lnOrd = reader.GetOrdinal(schema[2].ColumnName);
            ageOrd = reader.GetOrdinal(schema[3].ColumnName);
        }

        protected override bool ValidateCustom(DataValidationContext context, ExcelDataReader dr)
        {
            var valid = true;

            if (context.IsValid(idOrd)) 
            {
                var id = context.GetValue<int>(idOrd);
                if (!seenIds.Add(id))
                {
                    // already seen it
                    Console.WriteLine($"Duplicate Id at row {dr.RowNumber} col {idOrd} value '{id}'.");
                }
            }

            var fn = dr.GetString(fnOrd);
            if (string.IsNullOrEmpty(fn))
            {
                LogError(dr, fnOrd);
                valid = false;
            }

            var ln = dr.GetString(lnOrd);
            if (string.IsNullOrEmpty(ln))
            {
                LogError(dr, lnOrd);
                valid = false;
            }

            var age = context.GetValue<int>(ageOrd);
            if (age < 0)
            {
                LogError(dr, ageOrd);
                valid = false;
            }

            return valid;
        }
    }

    class EmailValidator : Validator
    {
        int idOrd, userOrd, emailOrd;

        public EmailValidator(DbDataReader reader, Schema schema)
        {
            idOrd = reader.GetOrdinal(schema[0].ColumnName);
            userOrd = reader.GetOrdinal(schema[1].ColumnName);
            emailOrd = reader.GetOrdinal(schema[2].ColumnName);
        }

        protected override bool ValidateCustom(DataValidationContext context, ExcelDataReader edr)
        {
            var valid = true;
            var email = edr.GetString(emailOrd);
            if (!email.Contains("@"))
            {
                LogError(edr, emailOrd);
                valid = false;
            }

            return valid;
        }
    }

}

static class Ex
{
    // I will add an efficient version of this API to Sylvan.Data
    public static bool IsValid(this DataValidationContext c, int ord)
    {
        // this is a bit slower than it needs to be
        return c.GetErrors().Contains(ord) == false;
    }
}