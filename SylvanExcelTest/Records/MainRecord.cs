namespace SylvanExcelTest.Records;

public record MainRecord
{
    public List<UserRecord>? Users { get; set; }
    public List<EmailRecord>? Emails { get; set; }
}