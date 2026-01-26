using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Application.UseCases;

             
                                                          
              
public class DemoOperationUseCase
{
    private readonly IWorkerService _workerService;

    public DemoOperationUseCase(IWorkerService workerService)
    {
        _workerService = workerService;
    }

                 
                                 
                  
                                                                       
                                                                     
                                                                  
                                                                                 
                                                     
                                                               
                                                                           
                                                                       
    public async Task<Result<DemoOperationResponse>> ExecuteAsync(
        int durationSeconds = 10,
        int itemCount = 10,
        bool simulateError = false,
        int errorAtPercent = 50,
        Action<LogLevel, string>? onLog = null,
        Action<int, string?>? onProgress = null,
        Action<DemoItemResult>? onPartialOutput = null,
        CancellationToken cancellationToken = default)
    {
        var request = new DemoOperationRequest
        {
            DurationSeconds = durationSeconds,
            ItemCount = itemCount,
            SimulateError = simulateError,
            ErrorAtPercent = errorAtPercent
        };

        return await _workerService.RunDemoOperationAsync(
            request,
            evt =>
            {
                switch (evt.EventType)
                {
                    case EventType.Log:
                        var logPayload = JsonMessageSerializer.ExtractPayload<LogEventPayload>(evt.Payload);
                        if (logPayload != null)
                        {
                            onLog?.Invoke(logPayload.Level, logPayload.Message);
                        }
                        break;

                    case EventType.Progress:
                        var progressPayload = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(evt.Payload);
                        if (progressPayload != null)
                        {
                            onProgress?.Invoke(progressPayload.PercentComplete, progressPayload.StatusMessage);
                        }
                        break;

                    case EventType.PartialOutput:
                        var partialPayload = JsonMessageSerializer.ExtractPayload<PartialOutputPayload>(evt.Payload);
                        if (partialPayload?.Data != null)
                        {
                            var item = JsonMessageSerializer.ExtractPayload<DemoItemResult>(partialPayload.Data);
                            if (item != null)
                            {
                                onPartialOutput?.Invoke(item);
                            }
                        }
                        break;
                }
            },
            cancellationToken);
    }
}
