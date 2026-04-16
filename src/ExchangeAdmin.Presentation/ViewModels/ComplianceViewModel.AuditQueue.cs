using System.ComponentModel;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Localization;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed partial class ComplianceViewModel
{
    public string? AuditWarning => SelectedAuditSearchTask?.Warning;
    public bool HasAuditWarning => !string.IsNullOrWhiteSpace(AuditWarning);

    public string? AuditErrorMessage => SelectedAuditSearchTask?.ErrorMessage;
    public bool HasAuditError => !string.IsNullOrWhiteSpace(AuditErrorMessage);

    public bool HasSelectedAuditSearchTask => SelectedAuditSearchTask != null;

    public string? SelectedAuditTaskReferenceText => !string.IsNullOrWhiteSpace(SelectedAuditSearchTask?.CorrelationId)
        ? $"corr={SelectedAuditSearchTask!.CorrelationId}"
        : null;

    public bool HasSelectedAuditTaskReference => !string.IsNullOrWhiteSpace(SelectedAuditTaskReferenceText);

    public string? SelectedAuditTaskFilterSummary => SelectedAuditSearchTask?.FilterSummary;

    public string? SelectedAuditTaskStatusMessage => SelectedAuditSearchTask?.StatusMessage;

    private Task EnqueueAuditSearchAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearStateForDisconnectedSession();
            return Task.CompletedTask;
        }

        if (AuditEndDate.Date < AuditStartDate.Date)
        {
            ErrorMessage = Loc.Get("Compliance.AuditDateRangeError");
            return Task.CompletedTask;
        }

        ErrorMessage = null;

        var task = new ComplianceAuditSearchTaskViewModel(
            _nextAuditTaskSequence++,
            new SearchUnifiedAuditLogRequest
            {
                StartDate = AuditStartDate.Date,
                EndDate = AuditEndDate.Date.AddDays(1).AddTicks(-1),
                UserIds = ParseMultiValue(AuditUserIdsText),
                Operations = ParseMultiValue(AuditOperationsText),
                ObjectIds = ParseMultiValue(AuditObjectIdsText),
                FreeText = string.IsNullOrWhiteSpace(AuditFreeText) ? null : AuditFreeText.Trim(),
                MaxResults = AuditMaxResults
            });

        RegisterAuditSearchTask(task);
        SelectedAuditSearchTask = task;
        _shellViewModel.AddLog(LogLevel.Information, $"{task.Title} queued for Search-UnifiedAuditLog.", "Compliance");

        EnsureAuditQueueProcessorRunning();
        return Task.CompletedTask;
    }

    private string GetAuditStatusText()
    {
        return SelectedAuditSearchTask == null
            ? (AuditSearchTasks.Count == 0
                ? Loc.Get("Compliance.AuditNoJobsQueued")
                : Loc.GetFormat("Compliance.AuditJobCount", AuditSearchTasks.Count))
            : Loc.GetFormat(
                "Compliance.AuditStatusFormat",
                SelectedAuditSearchTask.Title,
                SelectedAuditSearchTask.StateLabel,
                SelectedAuditSearchTask.ResultCount);
    }

    private string GetAuditTaskCenterStatus()
    {
        if (AuditSearchTasks.Count == 0)
        {
            return Loc.Get("Compliance.AuditNoJobsQueued");
        }

        var runningCount = AuditSearchTasks.Count(task => task.IsRunning);
        var queuedCount = AuditSearchTasks.Count(task => task.IsQueued);
        var completedCount = AuditSearchTasks.Count(task => task.IsCompleted);
        var failedCount = AuditSearchTasks.Count(task => task.IsFailed);

        var parts = new List<string> { Loc.GetFormat("Compliance.AuditTotal", AuditSearchTasks.Count) };
        if (runningCount > 0)
        {
            parts.Add(Loc.GetFormat("Compliance.AuditRunning", runningCount));
        }

        if (queuedCount > 0)
        {
            parts.Add(Loc.GetFormat("Compliance.AuditQueued", queuedCount));
        }

        if (completedCount > 0)
        {
            parts.Add(Loc.GetFormat("Compliance.AuditCompleted", completedCount));
        }

        if (failedCount > 0)
        {
            parts.Add(Loc.GetFormat("Compliance.AuditFailed", failedCount));
        }

        return string.Join(" | ", parts);
    }

    private void RegisterAuditSearchTask(ComplianceAuditSearchTaskViewModel task)
    {
        task.PropertyChanged += OnAuditSearchTaskPropertyChanged;
        AuditSearchTasks.Add(task);
        RaiseAuditTaskCollectionPropertiesChanged();
    }

    private void EnsureAuditQueueProcessorRunning()
    {
        if (_isAuditQueueProcessing)
        {
            return;
        }

        _isAuditQueueProcessing = true;
        _ = ProcessAuditQueueAsync();
    }

    private async Task ProcessAuditQueueAsync()
    {
        try
        {
            while (true)
            {
                var nextTask = AuditSearchTasks.FirstOrDefault(task => task.IsQueued);
                if (nextTask == null)
                {
                    break;
                }

                if (!_shellViewModel.IsExchangeConnected)
                {
                    nextTask.MarkFailed("Exchange session disconnected before execution.", nextTask.CorrelationId);
                    continue;
                }

                await ExecuteAuditSearchTaskAsync(nextTask);
            }
        }
        finally
        {
            _isAuditQueueProcessing = false;
            RefreshAuditQueueFlags();

            if (AuditSearchTasks.Any(task => task.IsQueued))
            {
                EnsureAuditQueueProcessorRunning();
            }
        }
    }

    private async Task ExecuteAuditSearchTaskAsync(ComplianceAuditSearchTaskViewModel task)
    {
        task.MarkRunning();
        RefreshAuditQueueFlags();

        try
        {
            var result = await _workerService.SearchUnifiedAuditLogAsync(
                task.Request,
                eventHandler: evt => HandleAuditTaskEvent(task, evt),
                cancellationToken: CancellationToken.None);

            var correlationId = task.CorrelationId ?? result.CorrelationId;
            task.CorrelationId = correlationId;

            if (!result.IsSuccess || result.Value == null)
            {
                var message = result.Error?.Message ?? "Audit search failed.";
                task.MarkFailed(message, correlationId);
                _shellViewModel.AddLog(LogLevel.Error, $"{task.Title} failed: {message}", "Compliance", correlationId);
                return;
            }

            task.MarkCompleted(result.Value, correlationId);

            if (task.HasWarning)
            {
                _shellViewModel.AddLog(LogLevel.Warning, $"{task.Title}: {task.Warning}", "Compliance", correlationId);
            }
            else
            {
                _shellViewModel.AddLog(LogLevel.Information, $"{task.Title} completed with {task.ResultCount} records.", "Compliance", correlationId);
            }
        }
        catch (Exception ex)
        {
            var correlationId = task.CorrelationId;
            task.MarkFailed(ex.Message, correlationId);
            _shellViewModel.AddLog(LogLevel.Error, $"{task.Title} failed: {ex.Message}", "Compliance", correlationId);
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(task.CorrelationId))
            {
                _shellViewModel.UnregisterBackgroundProgressOperation(task.CorrelationId);
            }

            RefreshAuditQueueFlags();
            if (ReferenceEquals(SelectedAuditSearchTask, task))
            {
                SyncSelectedAuditTaskState();
            }
        }
    }

    private void HandleAuditTaskEvent(ComplianceAuditSearchTaskViewModel task, EventEnvelope evt)
    {
        RunOnUiThread(() =>
        {
            if (!string.IsNullOrWhiteSpace(evt.CorrelationId))
            {
                task.CorrelationId = evt.CorrelationId;
                _shellViewModel.RegisterBackgroundProgressOperation(evt.CorrelationId);
            }

            if (evt.EventType != EventType.Progress)
            {
                return;
            }

            var progress = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(evt.Payload);
            if (progress == null)
            {
                return;
            }

            task.ApplyProgress(progress);
            RefreshAuditQueueFlags();

            if (ReferenceEquals(SelectedAuditSearchTask, task))
            {
                SyncSelectedAuditTaskState();
            }
        });
    }

    private void RefreshAuditQueueFlags()
    {
        IsSearchingAudit = AuditSearchTasks.Any(task => task.IsRunning) || _isAuditQueueProcessing;
        RaiseAuditTaskCollectionPropertiesChanged();
    }

    private void SyncSelectedAuditTaskState()
    {
        AuditResults.ReplaceAll(SelectedAuditSearchTask?.Results ?? Array.Empty<UnifiedAuditLogRecordDto>());
        OnPropertyChanged(nameof(AuditStatusText));
        OnPropertyChanged(nameof(AuditWarning));
        OnPropertyChanged(nameof(HasAuditWarning));
        OnPropertyChanged(nameof(AuditErrorMessage));
        OnPropertyChanged(nameof(HasAuditError));
        OnPropertyChanged(nameof(SelectedAuditTaskReferenceText));
        OnPropertyChanged(nameof(HasSelectedAuditTaskReference));
        OnPropertyChanged(nameof(SelectedAuditTaskFilterSummary));
        OnPropertyChanged(nameof(SelectedAuditTaskStatusMessage));
    }

    private void RaiseAuditTaskCollectionPropertiesChanged()
    {
        OnPropertyChanged(nameof(HasAuditSearchTasks));
        OnPropertyChanged(nameof(AuditTaskCenterStatus));
        OnPropertyChanged(nameof(AuditStatusText));
    }

    private void OnAuditSearchTaskPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        RaiseAuditTaskCollectionPropertiesChanged();

        if (ReferenceEquals(sender, SelectedAuditSearchTask))
        {
            SyncSelectedAuditTaskState();
        }
    }

    partial void OnDisconnectedSessionAuditCleanup()
    {
        foreach (var auditTask in AuditSearchTasks.Where(task => task.IsQueued).ToList())
        {
            auditTask.MarkFailed("Exchange session disconnected before execution.", auditTask.CorrelationId);
        }

        if (SelectedAuditSearchTask == null && AuditSearchTasks.Count == 0)
        {
            AuditResults.Clear();
        }

        SyncSelectedAuditTaskState();
        RaiseAuditTaskCollectionPropertiesChanged();
    }
}

