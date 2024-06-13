
public record UserRecord
{
    public int Id { get; set; }
    public string? FirstName { get; set; }
    public string? LastName { get; set; }
    public int Age { get; set; }
}

public record EmailRecord
{
    public int Id { get; set; }
    public int UserId { get; set; }
    public string Email { get; set; }
}
