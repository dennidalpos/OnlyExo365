using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace ExchangeAdmin.Worker.PowerShell;

internal sealed class PowerShellExecutionPipeline
{
    private readonly PowerShellEngine _engine;
    private readonly RunspaceRecoveryService _runspaceRecoveryService;

    public PowerShellExecutionPipeline(PowerShellEngine engine, RunspaceRecoveryService runspaceRecoveryService)
    {
        _engine = engine;
        _runspaceRecoveryService = runspaceRecoveryService;
    }

    public async Task<PowerShellResult> ExecuteAsync(
        string script,
        Dictionary<string, object>? parameters = null,
        Action<string, string>? onVerbose = null,
        Action<string, string>? onWarning = null,
        Action<ErrorRecord>? onError = null,
        Action<PSObject>? onOutput = null,
        CancellationToken cancellationToken = default)
    {
        if (_engine.IsDisposingRequested)
        {
            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = "PowerShell engine is disposing"
            };
        }

        if (!_engine.Initialized || _engine.Runspace == null)
        {
            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = "PowerShell engine not initialized"
            };
        }

        if (!_runspaceRecoveryService.IsRunspaceUsable())
        {
            Debug.WriteLine("[PowerShellEngine] Runspace not usable, attempting recovery...");

            var recovered = await _runspaceRecoveryService.TryRecoverRunspaceAsync().ConfigureAwait(false);
            if (!recovered)
            {
                return new PowerShellResult
                {
                    Success = false,
                    ErrorMessage = "Runspace is in invalid state and recovery failed",
                    RunspaceCorrupted = true
                };
            }
        }

        await _engine.ExecutionLock.WaitAsync(cancellationToken).ConfigureAwait(false);

        System.Management.Automation.PowerShell? ps = null;

        try
        {
#pragma warning disable CA2000
            ps = System.Management.Automation.PowerShell.Create();
#pragma warning restore CA2000
            ps.Runspace = _engine.Runspace;
            ps.AddScript(script);

            if (parameters != null)
            {
                foreach (var parameter in parameters)
                {
                    ps.AddParameter(parameter.Key, parameter.Value);
                }
            }

            var verbose = new List<string>();
            var warnings = new List<string>();
            var errors = new List<ErrorRecord>();
            var output = new List<PSObject>();

            ps.Streams.Verbose.DataAdded += (_, e) =>
            {
                if (e.Index >= 0 && e.Index < ps.Streams.Verbose.Count)
                {
                    var message = ps.Streams.Verbose[e.Index].Message;
                    verbose.Add(message);
                    onVerbose?.Invoke("Verbose", message);
                }
            };

            ps.Streams.Warning.DataAdded += (_, e) =>
            {
                if (e.Index >= 0 && e.Index < ps.Streams.Warning.Count)
                {
                    var message = ps.Streams.Warning[e.Index].Message;
                    warnings.Add(message);
                    onWarning?.Invoke("Warning", message);
                }
            };

            ps.Streams.Error.DataAdded += (_, e) =>
            {
                if (e.Index < 0 || e.Index >= ps.Streams.Error.Count)
                {
                    return;
                }

                var record = ps.Streams.Error[e.Index];
                var errorMessage = record.Exception?.Message ?? record.ToString();
                if (errorMessage.Contains("deprecat", StringComparison.OrdinalIgnoreCase) ||
                    errorMessage.Contains("will start deprecating", StringComparison.OrdinalIgnoreCase))
                {
                    warnings.Add(errorMessage);
                    onWarning?.Invoke("Warning", errorMessage);
                    return;
                }

                errors.Add(record);
                onError?.Invoke(record);
            };

            var registration = cancellationToken.Register(() =>
            {
                try
                {
                    Debug.WriteLine("[PowerShellEngine] Cancellation requested, stopping pipeline...");
                    ps.Stop();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerShellEngine] Error stopping pipeline: {ex.Message}");
                }
            });

            try
            {
                ConsoleLogger.Debug("PowerShellEngine", $"Starting script execution (length: {script.Length} chars)");
                var scriptPreview = script.Length > 100 ? script[..100] + "..." : script;
                ConsoleLogger.Verbose("PowerShellEngine", $"Script preview (sanitized): {scriptPreview.Replace("\n", " ").Replace("\r", "")}");

                await Task.Run(() =>
                {
                    var collection = new PSDataCollection<PSObject>();
                    collection.DataAdded += (_, e) =>
                    {
                        if (e.Index >= 0 && e.Index < collection.Count)
                        {
                            var item = collection[e.Index];
                            output.Add(item);
                            onOutput?.Invoke(item);
                        }
                    };

                    ConsoleLogger.Debug("PowerShellEngine", "Invoking PowerShell...");
                    ps.Invoke(null, collection);
                    ConsoleLogger.Debug("PowerShellEngine", $"PowerShell invocation completed. HadErrors: {ps.HadErrors}, Output count: {output.Count}");
                }, cancellationToken).ConfigureAwait(false);

                if (cancellationToken.IsCancellationRequested)
                {
                    Debug.WriteLine("[PowerShellEngine] Execution cancelled");

                    return new PowerShellResult
                    {
                        Success = false,
                        WasCancelled = true,
                        Output = output,
                        Errors = errors,
                        Verbose = verbose,
                        Warning = warnings
                    };
                }

                var hasErrors = errors.Count > 0;
                if (!hasErrors)
                {
                    _engine.ConsecutiveFailures = 0;
                }

                return new PowerShellResult
                {
                    Success = !hasErrors,
                    Output = output,
                    Errors = errors,
                    Verbose = verbose,
                    Warning = warnings,
                    ErrorMessage = hasErrors ? string.Join("; ", errors.Select(e => e.ToString())) : null
                };
            }
            finally
            {
                await registration.DisposeAsync().ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            Debug.WriteLine("[PowerShellEngine] Operation cancelled");

            return new PowerShellResult
            {
                Success = false,
                WasCancelled = true
            };
        }
        catch (PSInvalidOperationException ex)
        {
            Debug.WriteLine($"[PowerShellEngine] PSInvalidOperationException: {ex.Message}");
            _engine.ConsecutiveFailures++;

            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                RunspaceCorrupted = _engine.ConsecutiveFailures >= PowerShellEngine.MaxConsecutiveFailuresBeforeReset
            };
        }
        catch (InvalidRunspaceStateException ex)
        {
            Debug.WriteLine($"[PowerShellEngine] InvalidRunspaceStateException: {ex.Message}");
            _engine.ConsecutiveFailures++;

            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                RunspaceCorrupted = true
            };
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PowerShellEngine] Execution error: {ex.GetType().Name} - {ex.Message}");
            _engine.ConsecutiveFailures++;

            return new PowerShellResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                RunspaceCorrupted = _engine.ConsecutiveFailures >= PowerShellEngine.MaxConsecutiveFailuresBeforeReset
            };
        }
        finally
        {
            if (ps != null)
            {
                try
                {
                    ps.Dispose();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PowerShellEngine] Error disposing PowerShell: {ex.Message}");
                }
            }

            _engine.ExecutionLock.Release();
        }
    }
}
