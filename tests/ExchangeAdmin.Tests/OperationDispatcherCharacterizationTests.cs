using System.Reflection;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Worker.Operations;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Tests;

public class OperationDispatcherCharacterizationTests
{
    [Fact]
    public void Constructor_RegistersEveryDeclaredOperationTypeExactlyOnce()
    {
        using var engine = new PowerShellEngine();
        var dispatcher = new OperationDispatcher(engine, _ => Task.CompletedTask);

        var registry = GetHandlerRegistry(dispatcher);
        var declaredOperations = Enum.GetValues<OperationType>();

        Assert.Equal(declaredOperations.Length, registry.Count);
        Assert.Equal(declaredOperations.OrderBy(static operation => operation).ToArray(),
            registry.Keys.OrderBy(static operation => operation).ToArray());
    }

    [Fact]
    public async Task DispatchAsync_ReturnsExplicitErrorForUnsupportedOperation()
    {
        using var engine = new PowerShellEngine();
        var dispatcher = new OperationDispatcher(engine, _ => Task.CompletedTask);

        var response = await dispatcher.DispatchAsync(new RequestEnvelope
        {
            CorrelationId = "corr-unsupported",
            Operation = (OperationType)int.MaxValue
        }, CancellationToken.None);

        Assert.False(response.Success);
        Assert.False(response.WasCancelled);
        Assert.NotNull(response.Error);
        Assert.Equal(ErrorCode.OperationNotSupported, response.Error!.Code);
        Assert.Contains("is not supported", response.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DispatchAsync_DoesNotConvertDeprecationTextInvalidOperationExceptionIntoSuccess()
    {
        using var engine = new PowerShellEngine();
        var dispatcher = new OperationDispatcher(engine, _ => Task.CompletedTask);
        var operation = OperationType.GetConnectionStatus;
        const string message = "This API will start deprecating soon, but the operation actually failed.";

        SetHandlerRegistry(dispatcher, new Dictionary<OperationType, IOperationAreaHandler>
        {
            [operation] = new ThrowingHandler(new InvalidOperationException(message))
        });

        var response = await dispatcher.DispatchAsync(new RequestEnvelope
        {
            CorrelationId = "corr-deprecation-failure",
            Operation = operation
        }, CancellationToken.None);

        Assert.False(response.Success);
        Assert.False(response.WasCancelled);
        Assert.NotNull(response.Error);
        Assert.Equal(ErrorCode.Unknown, response.Error!.Code);
        Assert.Equal(message, response.Error.Message);
    }

    private static IReadOnlyDictionary<OperationType, object> GetHandlerRegistry(OperationDispatcher dispatcher)
    {
        var registry = GetRegistryFieldValue(dispatcher);

        var typedRegistry = Assert.IsAssignableFrom<System.Collections.IDictionary>(registry);
        var mappedOperations = new Dictionary<OperationType, object>();

        foreach (System.Collections.DictionaryEntry entry in typedRegistry)
        {
            mappedOperations.Add(Assert.IsType<OperationType>(entry.Key), entry.Value!);
        }

        return mappedOperations;
    }

    private static void SetHandlerRegistry(OperationDispatcher dispatcher, IReadOnlyDictionary<OperationType, IOperationAreaHandler> registry)
    {
        var field = GetRegistryField();
        field.SetValue(dispatcher, registry);
    }

    private static object GetRegistryFieldValue(OperationDispatcher dispatcher)
    {
        var field = GetRegistryField();
        var registry = field.GetValue(dispatcher);
        Assert.NotNull(registry);
        return registry!;
    }

    private static FieldInfo GetRegistryField()
    {
        var field = typeof(OperationDispatcher).GetField("_handlerRegistry", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(field);
        return field!;
    }

    private sealed class ThrowingHandler(Exception exception) : IOperationAreaHandler
    {
        public IReadOnlyCollection<OperationType> SupportedOperations { get; } = [];

        public Task<ResponseEnvelope> HandleAsync(RequestEnvelope request, CancellationToken cancellationToken)
        {
            throw exception;
        }
    }
}
