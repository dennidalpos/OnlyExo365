using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class MailboxPermissionChangeItem
{
    public PermissionAction Action { get; init; }
    public PermissionType PermissionType { get; init; }
    public string User { get; init; } = string.Empty;
    public bool? AutoMapping { get; init; }
    public string Description { get; init; } = string.Empty;

    public PermissionDeltaActionDto ToDto()
    {
        return new PermissionDeltaActionDto
        {
            Action = Action,
            PermissionType = PermissionType,
            User = User,
            AutoMapping = AutoMapping,
            Description = Description
        };
    }
}
