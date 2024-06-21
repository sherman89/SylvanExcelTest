namespace SylvanExcelTest.Records;

public record UserRecord
{
    public int Id { get; set; }
    public string? FirstName { get; set; }
    public string? LastName { get; set; }
    public int Age { get; set; }
    public string? JobTitle { get; set; }
}