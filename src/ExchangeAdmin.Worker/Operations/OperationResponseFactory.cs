using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Worker.Operations;

internal sealed class OperationResponseFactory
{
    public ResponseEnvelope CreateSuccess<T>(string correlationId, T payload)
    {
        return new ResponseEnvelope
        {
            CorrelationId = correlationId,
            Success = true,
            Payload = JsonMessageSerializer.ToJsonElement(payload)
        };
    }

    public ResponseEnvelope CreateError(string correlationId, ErrorCode code, string message, bool isTransient = false, int? retryAfterSeconds = null)
    {
        return new ResponseEnvelope
        {
            CorrelationId = correlationId,
            Success = false,
            Error = new NormalizedErrorDto
            {
                Code = code,
                Message = message,
                IsTransient = isTransient,
                RetryAfterSeconds = retryAfterSeconds
            }
        };
    }

    public ResponseEnvelope CreateCancelled(string correlationId)
    {
        return new ResponseEnvelope
        {
            CorrelationId = correlationId,
            Success = false,
            WasCancelled = true
        };
    }
}
