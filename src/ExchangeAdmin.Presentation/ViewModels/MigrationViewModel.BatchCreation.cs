using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public partial class MigrationViewModel
{
    public IReadOnlyList<string> BatchTypes { get; } =
    [
        "Onboarding",
        "Offboarding",
        "IMAP"
    ];

    public string NewBatchName
    {
        get => _newBatchName;
        set
        {
            if (SetProperty(ref _newBatchName, value))
            {
                InvalidateBatchPreflight();
            }
        }
    }

    public string NewBatchType
    {
        get => _newBatchType;
        set
        {
            if (SetProperty(ref _newBatchType, NormalizeBatchType(value)))
            {
                InvalidateBatchPreflight();
            }
        }
    }

    public string? NewBatchEndpointIdentity
    {
        get => _newBatchEndpointIdentity;
        set
        {
            if (SetProperty(ref _newBatchEndpointIdentity, TrimToNull(value)))
            {
                InvalidateBatchPreflight(updateRecommendedBatchType: true);
            }
        }
    }

    public string NewBatchCsvFilePath
    {
        get => _newBatchCsvFilePath;
        set
        {
            if (SetProperty(ref _newBatchCsvFilePath, value))
            {
                InvalidateBatchPreflight();
            }
        }
    }

    public string? NewBatchTargetDeliveryDomain
    {
        get => _newBatchTargetDeliveryDomain;
        set
        {
            if (SetProperty(ref _newBatchTargetDeliveryDomain, value))
            {
                InvalidateBatchPreflight();
            }
        }
    }

    public string NewBatchNotificationEmailsText
    {
        get => _newBatchNotificationEmailsText;
        set => SetProperty(ref _newBatchNotificationEmailsText, value);
    }

    public bool NewBatchAutoStart
    {
        get => _newBatchAutoStart;
        set => SetProperty(ref _newBatchAutoStart, value);
    }

    public bool NewBatchAutoComplete
    {
        get => _newBatchAutoComplete;
        set => SetProperty(ref _newBatchAutoComplete, value);
    }

    public string? BatchPreflightSummary
    {
        get => _batchPreflightSummary;
        private set => SetProperty(ref _batchPreflightSummary, value);
    }

    public bool IsBatchPreflightReady
    {
        get => _isBatchPreflightReady;
        private set
        {
            if (SetProperty(ref _isBatchPreflightReady, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    private async Task RunBatchPreflightAsync(CancellationToken cancellationToken)
    {
        if (!CanRunBatchPreflight)
        {
            return;
        }

        IsRunningPreflight = true;
        ErrorMessage = null;
        BatchPreflightSummary = null;
        IsBatchPreflightReady = false;

        try
        {
            var result = await _workerService.GetMigrationBatchPreflightAsync(
                new GetMigrationBatchPreflightRequest
                {
                    Name = NewBatchName.Trim(),
                    BatchType = NormalizeBatchType(NewBatchType),
                    EndpointIdentity = NewBatchEndpointIdentity!.Trim(),
                    CsvFilePath = NewBatchCsvFilePath.Trim(),
                    TargetDeliveryDomain = TrimToNull(NewBatchTargetDeliveryDomain)
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = string.IsNullOrWhiteSpace(result.Error?.Message)
                    ? "Unable to run migration preflight."
                    : result.Error.Message;
                return;
            }

            IsBatchPreflightReady = result.Value.IsReady;
            BatchPreflightSummary = string.Join(
                Environment.NewLine,
                new[]
                {
                    $"Ready: {result.Value.IsReady}",
                    $"EndpointType: {result.Value.EndpointType ?? "(unknown)"}",
                    $"Csv rows: {result.Value.CsvRowCount}",
                    $"Headers: {(result.Value.CsvHeaders.Count == 0 ? "(none)" : string.Join(", ", result.Value.CsvHeaders))}",
                    string.Empty,
                    result.Value.Messages.Count == 0
                        ? "No preflight messages."
                        : string.Join(Environment.NewLine, result.Value.Messages)
                });

            _shellViewModel.AddLog(
                LogLevel.Information,
                $"Migration preflight completed: ready={result.Value.IsReady}, endpoint={NewBatchEndpointIdentity}",
                "Migration");
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsRunningPreflight = false;
        }
    }

    private async Task CreateBatchAsync(CancellationToken cancellationToken)
    {
        if (!CanCreateBatch)
        {
            return;
        }

        var createdName = NewBatchName.Trim();
        if (!ConfirmMutation(
                "Creating migration batch",
                createdName,
                "Create a new migration batch using the current endpoint, CSV, and options.",
                "Confirm migration batch creation"))
        {
            return;
        }

        IsCreatingBatch = true;
        ErrorMessage = null;
        var preservedEndpointIdentity = NewBatchEndpointIdentity;

        try
        {
            var result = await _workerService.CreateMigrationBatchAsync(
                new CreateMigrationBatchRequest
                {
                    Name = createdName,
                    BatchType = NormalizeBatchType(NewBatchType),
                    EndpointIdentity = NewBatchEndpointIdentity!.Trim(),
                    CsvFilePath = NewBatchCsvFilePath.Trim(),
                    TargetDeliveryDomain = TrimToNull(NewBatchTargetDeliveryDomain),
                    NotificationEmails = ParseNotificationEmails(NewBatchNotificationEmailsText),
                    AutoStart = NewBatchAutoStart,
                    AutoComplete = NewBatchAutoComplete
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to create the migration batch.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"Migration batch created: {createdName}", "Migration");
            await RefreshAsync(cancellationToken);

            var selected = Batches.FirstOrDefault(batch =>
                string.Equals(batch.Name, createdName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(batch.Identity, createdName, StringComparison.OrdinalIgnoreCase));

            if (selected != null)
            {
                SelectedBatch = selected;
            }

            ResetBatchCreationEditor(preservedEndpointIdentity);
            BatchPreflightSummary = $"Migration batch created: {createdName}";
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsCreatingBatch = false;
        }
    }

    private void ResetBatchCreationEditor(string? preservedEndpointIdentity = null)
    {
        NewBatchName = string.Empty;
        NewBatchType = "Onboarding";
        NewBatchEndpointIdentity = TrimToNull(preservedEndpointIdentity) ?? SelectedEndpoint?.Identity ?? Endpoints.FirstOrDefault()?.Identity;
        NewBatchCsvFilePath = string.Empty;
        NewBatchTargetDeliveryDomain = null;
        NewBatchNotificationEmailsText = string.Empty;
        NewBatchAutoStart = true;
        NewBatchAutoComplete = false;
        BatchPreflightSummary = null;
        IsBatchPreflightReady = false;
    }

    private void InvalidatePreflight()
    {
        BatchPreflightSummary = null;
        IsBatchPreflightReady = false;
    }

    private void UpdateRecommendedBatchType()
    {
        var endpoint = Endpoints.FirstOrDefault(candidate =>
            string.Equals(candidate.Identity, NewBatchEndpointIdentity, StringComparison.OrdinalIgnoreCase));

        if (endpoint == null)
        {
            return;
        }

        var recommendedType = NormalizeBatchType(NewBatchType);
        switch (NormalizeEndpointType(endpoint.EndpointType))
        {
            case "IMAP":
                recommendedType = "IMAP";
                break;
            case "ExchangeOutlookAnywhere":
                recommendedType = "Onboarding";
                break;
            case "ExchangeRemoteMove" when string.Equals(NewBatchType, "IMAP", StringComparison.OrdinalIgnoreCase):
                recommendedType = "Onboarding";
                break;
        }

        if (!string.Equals(NewBatchType, recommendedType, StringComparison.OrdinalIgnoreCase))
        {
            NewBatchType = recommendedType;
        }
    }

    private bool HasBatchPreflightMinimumData()
    {
        return !string.IsNullOrWhiteSpace(NewBatchName) &&
               !string.IsNullOrWhiteSpace(NewBatchEndpointIdentity) &&
               !string.IsNullOrWhiteSpace(NewBatchCsvFilePath);
    }

    private void InvalidateBatchPreflight(bool updateRecommendedBatchType = false)
    {
        InvalidatePreflight();

        if (updateRecommendedBatchType)
        {
            UpdateRecommendedBatchType();
        }

        RaiseCanExecuteChanged();
    }

    private static List<string> ParseNotificationEmails(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return [];
        }

        return value
            .Split([';', ',', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(static entry => !string.IsNullOrWhiteSpace(entry))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }
}
