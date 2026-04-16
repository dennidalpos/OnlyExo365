using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Worker.Operations;

public partial class OperationDispatcher
{
    private readonly PowerShellEngine _psEngine;
    private readonly CapabilityDetector _capabilityDetector;
    private readonly ExoCommands _exoCommands;
    private readonly ExoGroupCommands _exoGroupCommands;
    private readonly ExoPermissionCommands _exoPermissionCommands;
    private readonly WorkerConsoleController _consoleController;
    private readonly OperationEventPublisher _eventPublisher;
    private readonly OperationResponseFactory _responseFactory;
    private readonly IReadOnlyDictionary<OperationType, IOperationAreaHandler> _handlerRegistry;

    internal OperationDispatcher(
        PowerShellEngine psEngine,
        Func<EventEnvelope, Task> sendEvent,
        WorkerConsoleController? consoleController = null)
    {
        _psEngine = psEngine;
        _capabilityDetector = new CapabilityDetector(psEngine);
        _exoCommands = new ExoCommands(psEngine, _capabilityDetector);
        _exoGroupCommands = new ExoGroupCommands(psEngine, _capabilityDetector);
        _exoPermissionCommands = new ExoPermissionCommands(psEngine);
        _consoleController = consoleController ?? new WorkerConsoleController();
        _eventPublisher = new OperationEventPublisher(sendEvent);
        _responseFactory = new OperationResponseFactory();
        _handlerRegistry = CreateHandlerRegistry();
    }

    public async Task<ResponseEnvelope> DispatchAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        ConsoleLogger.Info("Dispatcher", $"Dispatching operation: {request.Operation}");
        try
        {
            if (_handlerRegistry.TryGetValue(request.Operation, out var handler))
            {
                return await handler.HandleAsync(request, cancellationToken);
            }

            return CreateErrorResponse(
                request.CorrelationId,
                ErrorCode.OperationNotSupported,
                $"Operation {request.Operation} is not supported");
        }
        catch (OperationCanceledException)
        {
            return CreateCancelledResponse(request.CorrelationId);
        }
        catch (Exception ex)
        {
            ConsoleLogger.Error("Dispatcher", $"Exception: {ex.GetType().Name} - {ex.Message}");
            ConsoleLogger.Verbose("Dispatcher", $"Stack trace: {ex.StackTrace}");
            var (code, isTransient, retryAfter) = ErrorClassifier.Classify(ex);
            await SendLogAsync(request.CorrelationId, LogLevel.Error, $"Operation failed: {ex.Message}");
            return CreateErrorResponse(request.CorrelationId, code, ex.Message, isTransient, retryAfter);
        }
    }

    private IReadOnlyDictionary<OperationType, IOperationAreaHandler> CreateHandlerRegistry()
    {
        var handlers = new IOperationAreaHandler[]
        {
            new ConnectionOperationsHandler(this),
            new RecipientOperationsHandler(this),
            new MailboxOperationsHandler(this),
            new GroupOperationsHandler(this),
            new MailSecurityOperationsHandler(this),
            new MailFlowOperationsHandler(this),
            new ComplianceOperationsHandler(this),
            new SupportOperationsHandler(this)
        };

        var registry = new Dictionary<OperationType, IOperationAreaHandler>();

        foreach (var handler in handlers)
        {
            foreach (var operation in handler.SupportedOperations)
            {
                registry.Add(operation, handler);
            }
        }

        return registry;
    }
}
