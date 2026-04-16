using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public partial class MigrationViewModel
{
    public IReadOnlyList<string> EndpointTypes { get; } =
    [
        "ExchangeRemoteMove",
        "ExchangeOutlookAnywhere",
        "IMAP"
    ];

    public IReadOnlyList<string> EndpointSecurities { get; } =
    [
        "Ssl",
        "Tls",
        "None"
    ];

    public IReadOnlyList<string> EndpointAuthentications { get; } =
    [
        "Basic",
        "Ntlm"
    ];

    public MigrationEndpointDto? SelectedEndpoint
    {
        get => _selectedEndpoint;
        set
        {
            if (SetProperty(ref _selectedEndpoint, value))
            {
                ApplyEndpointToEditor(value);
                if (value != null)
                {
                    NewBatchEndpointIdentity = value.Identity;
                }

                OnPropertyChanged(nameof(IsExistingEndpointSelected));
                RaiseCanExecuteChanged();
            }
        }
    }

    public string EndpointName
    {
        get => _endpointName;
        set
        {
            if (SetProperty(ref _endpointName, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public string EndpointType
    {
        get => _endpointType;
        set
        {
            if (SetProperty(ref _endpointType, NormalizeEndpointType(value)))
            {
                InvalidateEndpointEditor(refreshTypeFlags: true);
            }
        }
    }

    public string? EndpointRemoteServer
    {
        get => _endpointRemoteServer;
        set
        {
            if (SetProperty(ref _endpointRemoteServer, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public string? EndpointRpcProxyServer
    {
        get => _endpointRpcProxyServer;
        set
        {
            if (SetProperty(ref _endpointRpcProxyServer, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public string? EndpointExchangeServer
    {
        get => _endpointExchangeServer;
        set
        {
            if (SetProperty(ref _endpointExchangeServer, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public string? EndpointEmailAddress
    {
        get => _endpointEmailAddress;
        set
        {
            if (SetProperty(ref _endpointEmailAddress, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public string? EndpointRemoteTenant
    {
        get => _endpointRemoteTenant;
        set
        {
            if (SetProperty(ref _endpointRemoteTenant, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public int? EndpointPort
    {
        get => _endpointPort;
        set
        {
            if (SetProperty(ref _endpointPort, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public string EndpointSecurity
    {
        get => _endpointSecurity;
        set
        {
            if (SetProperty(ref _endpointSecurity, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public string EndpointAuthentication
    {
        get => _endpointAuthentication;
        set
        {
            if (SetProperty(ref _endpointAuthentication, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public string? EndpointUsername
    {
        get => _endpointUsername;
        set
        {
            if (SetProperty(ref _endpointUsername, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public bool HasEndpointPassword => !string.IsNullOrWhiteSpace(_pendingEndpointPassword);

    public int EndpointPasswordClearTrigger
    {
        get => _endpointPasswordClearTrigger;
        private set => SetProperty(ref _endpointPasswordClearTrigger, value);
    }

    public int? EndpointMaxConcurrentMigrations
    {
        get => _endpointMaxConcurrentMigrations;
        set
        {
            if (SetProperty(ref _endpointMaxConcurrentMigrations, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public int? EndpointMaxConcurrentIncrementalSyncs
    {
        get => _endpointMaxConcurrentIncrementalSyncs;
        set
        {
            if (SetProperty(ref _endpointMaxConcurrentIncrementalSyncs, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool EndpointSkipVerification
    {
        get => _endpointSkipVerification;
        set
        {
            if (SetProperty(ref _endpointSkipVerification, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public bool EndpointAcceptUntrustedCertificates
    {
        get => _endpointAcceptUntrustedCertificates;
        set
        {
            if (SetProperty(ref _endpointAcceptUntrustedCertificates, value))
            {
                InvalidateEndpointEditor();
            }
        }
    }

    public string? EndpointTestSummary
    {
        get => _endpointTestSummary;
        private set => SetProperty(ref _endpointTestSummary, value);
    }

    public bool IsExistingEndpointSelected => SelectedEndpoint != null;
    public bool IsImapEndpointType => string.Equals(EndpointType, "IMAP", StringComparison.OrdinalIgnoreCase);
    public bool IsExchangeRemoteMoveEndpointType => string.Equals(EndpointType, "ExchangeRemoteMove", StringComparison.OrdinalIgnoreCase);
    public bool IsExchangeOutlookAnywhereEndpointType => string.Equals(EndpointType, "ExchangeOutlookAnywhere", StringComparison.OrdinalIgnoreCase);

    public void SetEndpointPassword(string? value)
    {
        var normalized = string.IsNullOrEmpty(value) ? null : value;
        if (string.Equals(_pendingEndpointPassword, normalized, StringComparison.Ordinal))
        {
            return;
        }

        _pendingEndpointPassword = normalized;
        OnPropertyChanged(nameof(HasEndpointPassword));
        EndpointTestSummary = null;
        RaiseCanExecuteChanged();
    }

    private async Task RefreshEndpointsAsync(CancellationToken cancellationToken)
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearStateForDisconnectedSession();
            ErrorMessage = "Not connected to Exchange Online";
            return;
        }

        IsLoadingEndpoints = true;
        ErrorMessage = null;
        var previousIdentity = SelectedEndpoint?.Identity;

        try
        {
            var result = await _workerService.GetMigrationEndpointsAsync(
                new GetMigrationEndpointsRequest { SortBy = "Name" },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load migration endpoint";
                return;
            }

            Endpoints.ReplaceAll(result.Value.Endpoints);

            if (!string.IsNullOrWhiteSpace(previousIdentity))
            {
                SelectedEndpoint = Endpoints.FirstOrDefault(endpoint =>
                    string.Equals(endpoint.Identity, previousIdentity, StringComparison.OrdinalIgnoreCase));
            }
            else if (SelectedEndpoint != null)
            {
                SelectedEndpoint = null;
            }

            if (SelectedEndpoint == null && string.IsNullOrWhiteSpace(NewBatchEndpointIdentity) && Endpoints.Count > 0)
            {
                NewBatchEndpointIdentity = Endpoints[0].Identity;
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsLoadingEndpoints = false;
        }
    }

    private async Task SaveEndpointAsync(CancellationToken cancellationToken)
    {
        if (!CanSaveEndpoint)
        {
            return;
        }

        IsSavingEndpoint = true;
        ErrorMessage = null;
        EndpointTestSummary = null;
        var expectedIdentity = SelectedEndpoint?.Identity ?? EndpointName.Trim();
        var request = BuildUpsertEndpointRequest();

        if (!ConfirmMutation(
                "Saving migration endpoint",
                expectedIdentity,
                "Create or update the selected migration endpoint.",
                "Confirm endpoint save"))
        {
            return;
        }

        ClearEndpointPassword();

        try
        {
            var result = await _workerService.UpsertMigrationEndpointAsync(
                request,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to save the migration endpoint.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"Migration endpoint saved: {expectedIdentity}", "Migration");
            await RefreshEndpointsAsync(cancellationToken);

            var selected = Endpoints.FirstOrDefault(endpoint =>
                string.Equals(endpoint.Identity, expectedIdentity, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(endpoint.Name, expectedIdentity, StringComparison.OrdinalIgnoreCase));

            if (selected != null)
            {
                SelectedEndpoint = selected;
            }
            else
            {
                NewBatchEndpointIdentity = expectedIdentity;
            }
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsSavingEndpoint = false;
        }
    }

    private async Task TestEndpointAsync(CancellationToken cancellationToken)
    {
        if (!CanTestEndpoint)
        {
            return;
        }

        IsTestingEndpoint = true;
        ErrorMessage = null;
        EndpointTestSummary = null;
        var request = BuildTestEndpointRequest();
        ClearEndpointPassword();

        try
        {
            var result = await _workerService.TestMigrationEndpointAsync(
                request,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to test the migration endpoint.";
                return;
            }

            EndpointTestSummary = string.IsNullOrWhiteSpace(result.Value.Details)
                ? result.Value.Summary
                : $"{result.Value.Summary}{Environment.NewLine}{Environment.NewLine}{result.Value.Details?.Trim()}";

            _shellViewModel.AddLog(LogLevel.Information, "Migration endpoint test completed", "Migration");
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsTestingEndpoint = false;
        }
    }

    private UpsertMigrationEndpointRequest BuildUpsertEndpointRequest()
    {
        var password = CaptureEndpointPassword();
        return new UpsertMigrationEndpointRequest
        {
            Identity = SelectedEndpoint?.Identity,
            Name = EndpointName.Trim(),
            EndpointType = NormalizeEndpointType(EndpointType),
            RemoteServer = TrimToNull(EndpointRemoteServer),
            RpcProxyServer = TrimToNull(EndpointRpcProxyServer),
            ExchangeServer = TrimToNull(EndpointExchangeServer),
            EmailAddress = TrimToNull(EndpointEmailAddress),
            RemoteTenant = TrimToNull(EndpointRemoteTenant),
            Port = EndpointPort,
            Security = TrimToNull(EndpointSecurity),
            Authentication = TrimToNull(EndpointAuthentication),
            Username = TrimToNull(EndpointUsername),
            Password = password,
            MaxConcurrentMigrations = EndpointMaxConcurrentMigrations,
            MaxConcurrentIncrementalSyncs = EndpointMaxConcurrentIncrementalSyncs,
            SkipVerification = EndpointSkipVerification,
            AcceptUntrustedCertificates = EndpointAcceptUntrustedCertificates
        };
    }

    private TestMigrationEndpointRequest BuildTestEndpointRequest()
    {
        var useExistingEndpoint =
            SelectedEndpoint != null &&
            string.IsNullOrWhiteSpace(EndpointUsername) &&
            !HasEndpointPassword;

        var password = CaptureEndpointPassword();

        return new TestMigrationEndpointRequest
        {
            Identity = SelectedEndpoint?.Identity,
            UseExistingEndpoint = useExistingEndpoint,
            EndpointType = NormalizeEndpointType(EndpointType),
            RemoteServer = TrimToNull(EndpointRemoteServer),
            RpcProxyServer = TrimToNull(EndpointRpcProxyServer),
            ExchangeServer = TrimToNull(EndpointExchangeServer),
            EmailAddress = TrimToNull(EndpointEmailAddress),
            Port = EndpointPort,
            Security = TrimToNull(EndpointSecurity),
            Authentication = TrimToNull(EndpointAuthentication),
            Username = TrimToNull(EndpointUsername),
            Password = password,
            SkipVerification = EndpointSkipVerification,
            AcceptUntrustedCertificates = EndpointAcceptUntrustedCertificates
        };
    }

    private void ApplyEndpointToEditor(MigrationEndpointDto? endpoint)
    {
        if (endpoint == null)
        {
            EndpointName = string.Empty;
            EndpointType = "ExchangeRemoteMove";
            EndpointRemoteServer = null;
            EndpointRpcProxyServer = null;
            EndpointExchangeServer = null;
            EndpointEmailAddress = null;
            EndpointRemoteTenant = null;
            EndpointPort = 993;
            EndpointSecurity = "Ssl";
            EndpointAuthentication = "Basic";
            EndpointUsername = null;
            ClearEndpointPassword();
            EndpointMaxConcurrentMigrations = 20;
            EndpointMaxConcurrentIncrementalSyncs = 10;
            EndpointSkipVerification = false;
            EndpointAcceptUntrustedCertificates = false;
            EndpointTestSummary = null;
            return;
        }

        EndpointName = endpoint.Name;
        EndpointType = NormalizeEndpointType(endpoint.EndpointType);
        EndpointRemoteServer = endpoint.RemoteServer;
        EndpointRpcProxyServer = endpoint.RpcProxyServer;
        EndpointExchangeServer = endpoint.ExchangeServer;
        EndpointEmailAddress = endpoint.EmailAddress;
        EndpointRemoteTenant = endpoint.RemoteTenant;
        EndpointPort = endpoint.Port ?? (string.Equals(endpoint.EndpointType, "IMAP", StringComparison.OrdinalIgnoreCase) ? 993 : null);
        EndpointSecurity = TrimToNull(endpoint.Security) ?? "Ssl";
        EndpointAuthentication = TrimToNull(endpoint.Authentication) ?? "Basic";
        EndpointUsername = null;
        ClearEndpointPassword();
        EndpointMaxConcurrentMigrations = endpoint.MaxConcurrentMigrations;
        EndpointMaxConcurrentIncrementalSyncs = endpoint.MaxConcurrentIncrementalSyncs;
        EndpointSkipVerification = endpoint.SkipVerification ?? false;
        EndpointAcceptUntrustedCertificates = endpoint.AcceptUntrustedCertificates ?? false;
        EndpointTestSummary = null;
    }

    private void ResetEndpointEditor()
    {
        if (SetProperty(ref _selectedEndpoint, null, nameof(SelectedEndpoint)))
        {
            OnPropertyChanged(nameof(IsExistingEndpointSelected));
        }

        ApplyEndpointToEditor(null);
        RaiseCanExecuteChanged();
    }

    private bool HasEndpointEditorMinimumData()
    {
        if (string.IsNullOrWhiteSpace(EndpointName))
        {
            return false;
        }

        return NormalizeEndpointType(EndpointType) switch
        {
            "IMAP" => !string.IsNullOrWhiteSpace(EndpointRemoteServer) && EndpointPort is > 0,
            "ExchangeOutlookAnywhere" =>
                (!string.IsNullOrWhiteSpace(EndpointRemoteServer) ||
                 !string.IsNullOrWhiteSpace(EndpointRpcProxyServer) ||
                 !string.IsNullOrWhiteSpace(EndpointExchangeServer)) &&
                (IsExistingEndpointSelected || !RequiresCredential(NormalizeEndpointType(EndpointType)) || (!string.IsNullOrWhiteSpace(EndpointUsername) && HasEndpointPassword)),
            _ =>
                !string.IsNullOrWhiteSpace(EndpointRemoteServer) &&
                (IsExistingEndpointSelected || !RequiresCredential(NormalizeEndpointType(EndpointType)) || (!string.IsNullOrWhiteSpace(EndpointUsername) && HasEndpointPassword))
        };
    }

    private bool CanTestCurrentEndpoint()
    {
        return (SelectedEndpoint != null && string.IsNullOrWhiteSpace(EndpointUsername) && !HasEndpointPassword)
            || HasEndpointEditorMinimumData();
    }

    private void InvalidateEndpointEditor(bool refreshTypeFlags = false)
    {
        EndpointTestSummary = null;

        if (refreshTypeFlags)
        {
            OnPropertyChanged(nameof(IsImapEndpointType));
            OnPropertyChanged(nameof(IsExchangeRemoteMoveEndpointType));
            OnPropertyChanged(nameof(IsExchangeOutlookAnywhereEndpointType));

            if (IsImapEndpointType && EndpointPort is null or <= 0)
            {
                _endpointPort = 993;
                OnPropertyChanged(nameof(EndpointPort));
            }
        }

        RaiseCanExecuteChanged();
    }

    private string? CaptureEndpointPassword()
    {
        return string.IsNullOrWhiteSpace(_pendingEndpointPassword)
            ? null
            : _pendingEndpointPassword.Trim();
    }

    private void ClearEndpointPassword()
    {
        var hadValue = !string.IsNullOrWhiteSpace(_pendingEndpointPassword);
        _pendingEndpointPassword = null;
        EndpointPasswordClearTrigger++;

        if (hadValue)
        {
            OnPropertyChanged(nameof(HasEndpointPassword));
            EndpointTestSummary = null;
            RaiseCanExecuteChanged();
        }
    }
}


