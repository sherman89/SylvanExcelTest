using Sylvan.Data;

namespace SylvanExcelTest.Shared;

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