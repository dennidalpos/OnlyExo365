using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Application.UseCases;

public sealed class PrepareInteractiveExchangeSignInUseCase
{
    private readonly IInteractiveExchangeBootstrapService _bootstrapService;

    public PrepareInteractiveExchangeSignInUseCase(IInteractiveExchangeBootstrapService bootstrapService)
    {
        _bootstrapService = bootstrapService;
    }

    public Task<Result> ExecuteAsync(
        Action<LogLevel, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        return _bootstrapService.EnsureReadyAsync(onLog, cancellationToken);
    }
}
