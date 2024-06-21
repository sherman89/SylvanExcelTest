using SylvanExcelTest.Shared;

namespace SylvanExcelTest;

public class CodeLookupRepository
{
    public Dictionary<string, string> JobTitleFiCodeLookup { get; }
    public Dictionary<string, string> JobTitleEnCodeLookup { get; }

    public CodeLookupRepository()
    {
        JobTitleFiCodeLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Toimitusjohtaja", "1" },
            { "Johtava ohjelmistosuunnittelija", "3" }
        };

        JobTitleEnCodeLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "CEO", "1" },
            { "Principal Software Engineer", "5" }
        };
    }

    public string? GetJobTitleCode(Language language, string selite)
    {
        string? value;

        switch (language)
        {
            case Language.Finnish:
                if (!JobTitleFiCodeLookup.TryGetValue(selite, out value))
                {
                    JobTitleEnCodeLookup.TryGetValue(selite, out value);
                }
                break;
            case Language.English:
            default:
                if (!JobTitleEnCodeLookup.TryGetValue(selite, out value))
                {
                    JobTitleFiCodeLookup.TryGetValue(selite, out value);
                }
                break;
        }

        return value;
    }
}