using System.Collections;
using System.Management.Automation;
using System.Text.Json;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

internal abstract class ExoCommandModuleBase
{
    protected const string StructuredWarningPrefix = "__EA_WARN__";

    protected ExoCommandModuleBase(PowerShellEngine engine)
    {
        Engine = engine;
    }

    protected PowerShellEngine Engine { get; }

    protected Task<List<PSObject>> RunScriptAsync(string script, CancellationToken cancellationToken = default)
        => RunScriptAsync(script, null, cancellationToken);

    protected async Task<List<PSObject>> RunScriptAsync(
        string script,
        Dictionary<string, object>? parameters = null,
        CancellationToken cancellationToken = default)
    {
        var result = await Engine.ExecuteAsync(script, parameters, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        if (!result.Success)
        {
            throw new InvalidOperationException(result.ErrorMessage ?? "PowerShell script execution failed");
        }

        return result.Output;
    }

    protected async Task<List<PSObject>> RunScriptAllowErrorsAsync(
        string script,
        Dictionary<string, object>? parameters = null,
        CancellationToken cancellationToken = default)
    {
        var result = await Engine.ExecuteAsync(script, parameters, cancellationToken: cancellationToken);
        if (result.WasCancelled)
        {
            throw new OperationCanceledException();
        }

        return result.Output;
    }

    protected static object? GetPropertyValue(PSObject obj, string propertyName)
    {
        if (obj.Properties[propertyName] is { } property)
        {
            return property.Value;
        }

        if (obj.BaseObject is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                if (entry.Key is string key &&
                    string.Equals(key, propertyName, StringComparison.Ordinal))
                {
                    return entry.Value;
                }
            }
        }

        return null;
    }

    protected static string GetString(PSObject obj, string propertyName)
        => GetPropertyValue(obj, propertyName)?.ToString() ?? string.Empty;

    protected static string? GetNullableString(PSObject obj, string propertyName)
        => GetPropertyValue(obj, propertyName)?.ToString();

    protected static bool GetBool(PSObject obj, string propertyName)
    {
        var value = GetPropertyValue(obj, propertyName);
        if (value == null)
        {
            return false;
        }

        if (value is bool flag)
        {
            return flag;
        }

        return Convert.ToBoolean(value);
    }

    protected static DateTime? GetNullableDateTime(PSObject obj, string propertyName)
    {
        var value = GetPropertyValue(obj, propertyName);
        if (value == null)
        {
            return null;
        }

        if (value is DateTime timestamp)
        {
            return timestamp;
        }

        return DateTime.TryParse(value.ToString(), out var parsed)
            ? parsed
            : null;
    }

    protected static int? GetNullableInt(PSObject obj, string propertyName)
    {
        var value = GetPropertyValue(obj, propertyName);
        if (value == null)
        {
            return null;
        }

        return Convert.ToInt32(value);
    }

    protected static long? GetNullableLong(PSObject obj, string propertyName)
    {
        var value = GetPropertyValue(obj, propertyName);
        if (value == null)
        {
            return null;
        }

        if (value is long longValue)
        {
            return longValue;
        }

        return Convert.ToInt64(value);
    }

    protected static bool? GetNullableBool(PSObject obj, string propertyName)
    {
        var value = GetPropertyValue(obj, propertyName);
        if (value == null)
        {
            return null;
        }

        if (value is bool flag)
        {
            return flag;
        }

        return Convert.ToBoolean(value);
    }

    protected static List<string> ConvertToStringList(object? obj)
    {
        if (obj == null)
        {
            return new List<string>();
        }

        if (obj is object[] array)
        {
            return array
                .Select(x => x?.ToString() ?? string.Empty)
                .Where(x => !string.IsNullOrEmpty(x))
                .ToList();
        }

        if (obj is IEnumerable enumerable)
        {
            var list = new List<string>();
            foreach (var item in enumerable)
            {
                var str = item?.ToString();
                if (!string.IsNullOrEmpty(str))
                {
                    list.Add(str);
                }
            }

            return list;
        }

        var single = obj.ToString();
        return string.IsNullOrEmpty(single)
            ? new List<string>()
            : new List<string> { single };
    }

    protected static string EscapePs(string? value)
        => (value ?? string.Empty).Replace("'", "''");

    protected static string ToPsArrayLiteral(IEnumerable<string>? values)
    {
        if (values == null)
        {
            return "@()";
        }

        var normalized = values
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .Select(static value => $"'{EscapePs(value.Trim())}'")
            .ToArray();

        return normalized.Length == 0
            ? "@()"
            : "@(" + string.Join(", ", normalized) + ")";
    }

    protected static string ToPsBoolLiteral(bool value)
        => value ? "$true" : "$false";

    protected static List<OperationWarningDto> ParseStructuredWarnings(IEnumerable<string>? warnings)
    {
        var parsed = new List<OperationWarningDto>();
        if (warnings == null)
        {
            return parsed;
        }

        foreach (var warning in warnings)
        {
            if (string.IsNullOrWhiteSpace(warning))
            {
                continue;
            }

            if (warning.StartsWith(StructuredWarningPrefix, StringComparison.Ordinal))
            {
                var payload = warning[StructuredWarningPrefix.Length..];
                try
                {
                    var dto = JsonSerializer.Deserialize<OperationWarningDto>(payload);
                    if (dto != null && !string.IsNullOrWhiteSpace(dto.Message))
                    {
                        parsed.Add(dto);
                    }
                }
                catch
                {
                    parsed.Add(new OperationWarningDto
                    {
                        Code = "StructuredWarningParseFailed",
                        Scope = "Worker",
                        Message = $"Structured warning could not be parsed: {payload}",
                        IsPartialData = true
                    });
                }

                continue;
            }

            parsed.Add(new OperationWarningDto
            {
                Code = "PowerShellWarning",
                Scope = "Worker",
                Message = warning,
                IsPartialData = false
            });
        }

        return parsed;
    }

    protected static List<string> ExtractWarningMessages(IEnumerable<OperationWarningDto> warningDetails)
    {
        return warningDetails
            .Select(static warning => warning.Message?.Trim())
            .Where(static message => !string.IsNullOrWhiteSpace(message))
            .Distinct(StringComparer.Ordinal)
            .Cast<string>()
            .ToList();
    }

    protected async Task EnsureExchangeCmdletAvailableAsync(
        string cmdletName,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(cmdletName))
        {
            throw new InvalidOperationException("Cmdlet name is required.");
        }

        var (isConnected, _, _, _, _) = await Engine.GetConnectionStatusAsync(cancellationToken);
        Engine.Connected = isConnected;

        if (!isConnected)
        {
            throw new InvalidOperationException(
                $"Exchange Online session is not connected. The cmdlet {cmdletName} is unavailable until you reconnect.");
        }

        var validationResult = await ValidateExchangeCmdletAsync(cmdletName, cancellationToken);
        if (validationResult.Success)
        {
            return;
        }

        var (isStillConnected, _, _, _, _) = await Engine.GetConnectionStatusAsync(cancellationToken);
        Engine.Connected = isStillConnected;

        if (!isStillConnected)
        {
            throw new InvalidOperationException(
                $"Exchange Online session is not connected. The cmdlet {cmdletName} is unavailable until you reconnect.");
        }

        throw new InvalidOperationException(
            $"The Exchange Online cmdlet {cmdletName} is not available in the current session. Disconnect and reconnect, then retry.");
    }

    private Task<PowerShellResult> ValidateExchangeCmdletAsync(
        string cmdletName,
        CancellationToken cancellationToken)
    {
        var escapedCmdletName = EscapePs(cmdletName);
        var validationScript = $@"
if (-not (Get-Command -Name '{escapedCmdletName}' -ErrorAction SilentlyContinue)) {{
    throw 'Cmdlet {escapedCmdletName} is not available in the current Exchange Online session.'
}}";

        return Engine.ExecuteAsync(validationScript, cancellationToken: cancellationToken);
    }
}
