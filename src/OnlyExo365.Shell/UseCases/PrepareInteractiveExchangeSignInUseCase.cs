using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Shell.UseCases;

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

