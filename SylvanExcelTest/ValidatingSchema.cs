using System.Collections;
using System.Data.Common;

namespace Sylvan.Data.Excel;

sealed class ValidatingSchema
    : ExcelSchemaProvider, IEnumerable<Table>
{
    Dictionary<string, Table> tableLookup;
    List<Table> tables;

    public ValidatingSchema()
    {
        this.tables = new List<Table>();
        this.tableLookup = new Dictionary<string, Table>(StringComparer.OrdinalIgnoreCase);
    }

    public Table? FindTable(string name)
    {
        return
            this.tableLookup.TryGetValue(name, out var table)
            ? table
            : null;
    }

    public override DbColumn? GetColumn(string sheetName, string? name, int ordinal)
    {
        if (!tableLookup.TryGetValue(sheetName, out var table))
            return null;

        return table.FindColumn(sheetName, name, ordinal);
    }

    public override bool HasHeaders(string sheetName)
    {
        return true;
    }

    public void Add(Table table)
    {
        foreach (var name in table.Names)
        {
            if (tableLookup.ContainsKey(name))
            {
                throw new ArgumentException();
            }
        }
        foreach (var name in table.Names)
        {
            this.tableLookup.Add(name, table);
        }
        this.tables.Add(table);
    }

    public IEnumerator<Table> GetEnumerator()
    {
        return this.tables.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
}

public abstract class Table : IReadOnlyList<Column>
{
    string name;
    string[] names;

    Dictionary<string, Column> columnLookup;
    List<Column> columns;

    public int Count => columns.Count;

    public IEnumerable<string> Names => names;

    public Column this[int index] => columns[index];

    public Table(string name, string[] names = null)
    {
        this.names = names ?? [name];
        this.columns = new List<Column>();
        this.columnLookup = new Dictionary<string, Column>(StringComparer.OrdinalIgnoreCase);
    }

    internal Column? FindColumn(string sheetName, string? columnName, int ordinal)
    {
        if (columnName == null) return null;

        if (!columnLookup.TryGetValue(columnName, out var col)) return null;

        return col;
    }

    public abstract TableValidator GetValidator(DbDataReader dr, object? state = null);

    private protected void AddColumn(Column col)
    {
        foreach (var alias in col.Aliases)
        {
            if (columnLookup.ContainsKey(alias))
            {
                throw new ArgumentException();
            }
        }
        foreach (var alias in col.Aliases)
        {
            columnLookup.Add(alias, col);
        }
        this.columns.Add(col);
    }

    public IEnumerator<Column> GetEnumerator()
    {
        return this.columns.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
}

public class Table<T> : Table where T : TableValidator, new()
{
    public Table(string name, string[] names) : base(name, names)
    {
    }

    public override TableValidator GetValidator(DbDataReader dr, object? state = null)
    {
        var v = new T();
        v.table = this;
        v.SetState(state);

        var count = this.Count;
        var ordinals = new int[count];
        var validators = new ColumnValidator?[count];

        // at this point the schema is already mapped
        // we just need to sort out the ordinals
        var s = dr.GetColumnSchema();

        Array.Fill(ordinals, -1);

        for (int i = 0; i < count; i++)
        {
            var col = this[i];
            var name = col.ColumnName;
            validators[i] = ((Func<T, ColumnValidator>?)col.validator)?.Invoke(v);
            for (int j = 0; j < s.Count; j++)
            {
                if (s[j].ColumnName == name)
                {
                    ordinals[i] = j;
                    break;
                }
            }
        }
        v.ordinals = ordinals;
        v.validators = validators;
        return v;
    }

    public Table<T> Add<TCol>(string colName, string[] names, Func<T, ColumnValidator> columnValidator = null)
    {
        var col = new Column(colName, names, typeof(TCol), columnValidator);
        base.AddColumn(col);
        return this;
    }
}

public delegate bool ColumnValidator(DataValidationContext context, int ordinal);

public class TableValidator
{
    // Not super happy with internal here
    // but this has to be constructed externally, where T is present
    internal int[] ordinals;
    internal ColumnValidator?[] validators;
    internal Table table;

    public virtual void SetState(object? o) { }
       

    public virtual void LogError(DataValidationContext context, int ord, Exception? ex = null)
    {
        // Up to the user to override and do something.
    }

    public virtual bool Validate(DataValidationContext context)
    {
        return
            ValidateSchema(context) &&
            ValidateCustom(context);
    }

    protected virtual bool ValidateCustom(DataValidationContext context)
    {
        var valid = true;
        var dr = context.DataReader;
        // validate the columns in the order specified by the schema
        // not necessarily the order they appear in the file.
        for (int i = 0; i < ordinals.Length; i++)
        {
            var validator = validators[i];
            var ord = ordinals[i];
            if (ord == -1 || validator == null)
            {
                continue;
            }
            if (context.IsValid(ord))
            {
                
                try
                {
                    var isValid = validator(context, ord);
                    if (!isValid)
                    {
                        LogError(context, ord);
                        valid = false;
                    }
                } 
                catch(Exception ex)
                {
                    LogError(context, ord, ex);
                    valid = false;
                }
            }
        }
        return valid;
    }

    bool ValidateSchema(DataValidationContext context)
    {
        var valid = true;
        foreach (var ord in context.GetErrors())
        {
            var ex = context.GetException(ord);
            LogError(context, ord, ex);
            valid = false;
        }
        return valid;
    }
}

public sealed class Column : DbColumn
{
    string name;
    string[] names;
    internal object? validator;

    public IEnumerable<string> Aliases => names;

    public Column(string name, string[] names, Type type, object? validator)
    {
        this.validator = validator;
        this.name = name;
        this.names = names ?? [name];
        this.ColumnName = name;
        var t = Nullable.GetUnderlyingType(type);
        this.AllowDBNull = type.IsByRef || t != null;
        var dataType = t ?? type;
        this.DataType = dataType;
        this.DataTypeName = dataType.Name;
    }
}
