using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Shell.UseCases;

             
                                               
              
public class ConnectExchangeUseCase
{
    private readonly IConnectionWorkerService _workerService;

    public ConnectExchangeUseCase(IConnectionWorkerService workerService)
    {
        _workerService = workerService;
    }

                 
                                                            
                  
                                                     
                                                                       
    public async Task<Result<ConnectionStatusDto>> ExecuteAsync(
        Action<LogLevel, string>? onLog = null,
        CancellationToken cancellationToken = default)
    {
        return await _workerService.ConnectExchangeAsync(
            evt =>
            {
                if (evt.EventType == EventType.Log)
                {
                    var logPayload = JsonMessageSerializer.ExtractPayload<LogEventPayload>(evt.Payload);
                    if (logPayload != null)
                    {
                        onLog?.Invoke(logPayload.Level, logPayload.Message);
                    }
                }
            },
            cancellationToken);
    }
}

