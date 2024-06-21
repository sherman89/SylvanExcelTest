using Sylvan.Data;

namespace SylvanExcelTest.Shared;

public class DefaultValidator(List<string> errors) : BaseValidator(errors)
{
    protected override bool ValidateCustom(DataValidationContext context)
    {
        return true;
    }
}