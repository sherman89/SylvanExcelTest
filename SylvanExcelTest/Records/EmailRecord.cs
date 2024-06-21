namespace SylvanExcelTest.Records;

public record EmailRecord
{
    public int Id { get; set; }
    public int UserId { get; set; }
    public string? Email { get; set; }
}