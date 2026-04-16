namespace ExchangeAdmin.Contracts.Messages;

public sealed class SetWorkerConsoleVisibilityResponse
{
    public bool IsVisible { get; set; }

    public string? Message { get; set; }
}
