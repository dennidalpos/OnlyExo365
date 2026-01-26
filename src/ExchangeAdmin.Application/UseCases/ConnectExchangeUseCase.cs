using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Application.UseCases;

             
                                               
              
public class ConnectExchangeUseCase
{
    private readonly IWorkerService _workerService;

    public ConnectExchangeUseCase(IWorkerService workerService)
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
