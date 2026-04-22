using System.Reflection;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public sealed class ContactsViewModelTests
{
    [Fact]
    public async Task SaveCommand_CreateMailUser_ClearsPasswordAfterDispatch()
    {
        var worker = new ContactsWorkerServiceStub();
        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        ErrorDialogService.ConfirmationHandlerOverride = (_, _) => true;
        var viewModel = new ContactsViewModel(worker, shell)
        {
            ContactKind = "MailUser",
            DisplayName = "Mario Rossi",
            Alias = "mrossi",
            PrimarySmtpAddress = "mario.rossi@contoso.com",
            ExternalEmailAddress = "mario.rossi@gmail.com",
            UserPrincipalName = "mario.rossi@contoso.com"
        };

        viewModel.SetMailUserPassword("Sup3rSecret!");

        try
        {
            viewModel.SaveCommand.Execute(null);
            await WaitForConditionAsync(() => worker.UpsertContactCalls == 1);

            Assert.Equal("Sup3rSecret!", worker.LastUpsertContactRequest?.Password);
            Assert.False(viewModel.HasMailUserPassword);
            Assert.True(viewModel.MailUserPasswordClearTrigger > 0);
        }
        finally
        {
            ErrorDialogService.ConfirmationHandlerOverride = null;
        }
    }

    [Fact]
    public void DisconnectExchange_ClearsPendingMailUserPassword()
    {
        var worker = new ContactsWorkerServiceStub();
        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new ContactsViewModel(worker, shell)
        {
            ContactKind = "MailUser",
            DisplayName = "Mario Rossi",
            Alias = "mrossi",
            PrimarySmtpAddress = "mario.rossi@contoso.com",
            ExternalEmailAddress = "mario.rossi@gmail.com",
            UserPrincipalName = "mario.rossi@contoso.com"
        };

        viewModel.SetMailUserPassword("Sup3rSecret!");

        SetExchangeDisconnected(shell);

        Assert.False(viewModel.HasMailUserPassword);
        Assert.True(viewModel.MailUserPasswordClearTrigger > 0);
    }

    [Fact]
    public async Task LoadAsync_UsesPageSize250ForContactsPagination()
    {
        var worker = new ContactsWorkerServiceStub();
        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new ContactsViewModel(worker, shell);

        await viewModel.LoadAsync();

        Assert.Single(worker.ContactRequests);
        Assert.Equal(250, worker.ContactRequests[0].PageSize);
    }

    [Fact]
    public async Task RefreshCommand_PreservesLoadedContactDepth()
    {
        var worker = new ContactsWorkerServiceStub();
        worker.ContactResponses.Enqueue(CreateContactsResponse(250, totalCount: 520, hasMore: true));
        worker.ContactResponses.Enqueue(CreateContactsResponse(250, skip: 250, totalCount: 520, hasMore: true));
        worker.ContactResponses.Enqueue(CreateContactsResponse(500, totalCount: 520, hasMore: true));
        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new ContactsViewModel(worker, shell);

        await viewModel.LoadAsync();
        viewModel.LoadMoreCommand.Execute(null);
        await WaitForConditionAsync(() => worker.ContactRequests.Count == 2);
        viewModel.RefreshCommand.Execute(null);
        await WaitForConditionAsync(() => worker.ContactRequests.Count == 3);

        Assert.Equal([250, 250, 500], worker.ContactRequests.Select(request => request.PageSize));
        Assert.Equal(500, viewModel.Contacts.Count);
    }

    [Fact]
    public async Task SelectingContact_UsesSelectedFallbackValuesWhenDetailPayloadIsSparse()
    {
        var worker = new ContactsWorkerServiceStub
        {
            ContactDetailsFactory = request => Task.FromResult(new ContactDetailsDto
            {
                Identity = request.Identity,
                ContactKind = string.Empty,
                DisplayName = string.Empty,
                Name = null,
                Alias = null,
                PrimarySmtpAddress = string.Empty,
                ExternalEmailAddress = null,
                UserPrincipalName = null,
                HiddenFromAddressListsEnabled = true
            })
        };
        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new ContactsViewModel(worker, shell);
        var selected = new ContactListItemDto
        {
            Identity = "contact-01",
            ContactKind = "MailUser",
            DisplayName = "Mario Rossi",
            Name = "Mario Rossi",
            Alias = "mrossi",
            PrimarySmtpAddress = "mario.rossi@contoso.com",
            ExternalEmailAddress = "mario.rossi@gmail.com",
            UserPrincipalName = "mario.rossi@contoso.com"
        };

        viewModel.SelectedContact = selected;

        await WaitForConditionAsync(() => worker.ContactDetailRequests.Count == 1);

        Assert.Equal(selected.Identity, viewModel.ContactIdentity);
        Assert.Equal(selected.ContactKind, viewModel.ContactKind);
        Assert.Equal(selected.DisplayName, viewModel.DisplayName);
        Assert.Equal(selected.Name, viewModel.Name);
        Assert.Equal(selected.Alias, viewModel.Alias);
        Assert.Equal(selected.PrimarySmtpAddress, viewModel.PrimarySmtpAddress);
        Assert.Equal(selected.ExternalEmailAddress, viewModel.ExternalEmailAddress);
        Assert.Equal(selected.UserPrincipalName, viewModel.UserPrincipalName);
        Assert.True(viewModel.HiddenFromAddressListsEnabled);
    }

    [Fact]
    public async Task SelectingSecondContact_IgnoresStaleDetailsFromPreviousSelection()
    {
        var firstRequest = new TaskCompletionSource<ContactDetailsDto>(TaskCreationOptions.RunContinuationsAsynchronously);
        var secondRequest = new TaskCompletionSource<ContactDetailsDto>(TaskCreationOptions.RunContinuationsAsynchronously);
        var worker = new ContactsWorkerServiceStub
        {
            ContactDetailsFactory = request =>
            {
                if (request.Identity == "contact-01")
                {
                    return firstRequest.Task;
                }

                return secondRequest.Task;
            }
        };
        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new ContactsViewModel(worker, shell);
        var first = new ContactListItemDto
        {
            Identity = "contact-01",
            ContactKind = "MailContact",
            DisplayName = "First Contact",
            PrimarySmtpAddress = "first@contoso.com"
        };
        var second = new ContactListItemDto
        {
            Identity = "contact-02",
            ContactKind = "MailUser",
            DisplayName = "Second Contact",
            Name = "Second Contact",
            Alias = "second",
            PrimarySmtpAddress = "second@contoso.com",
            ExternalEmailAddress = "second@gmail.com",
            UserPrincipalName = "second@contoso.com"
        };

        viewModel.SelectedContact = first;
        viewModel.SelectedContact = second;

        await WaitForConditionAsync(() => worker.ContactDetailRequests.Count == 2);

        secondRequest.SetResult(new ContactDetailsDto
        {
            Identity = second.Identity,
            ContactKind = second.ContactKind,
            DisplayName = second.DisplayName,
            Name = second.Name,
            Alias = second.Alias,
            PrimarySmtpAddress = second.PrimarySmtpAddress,
            ExternalEmailAddress = second.ExternalEmailAddress,
            UserPrincipalName = second.UserPrincipalName
        });
        await WaitForConditionAsync(() => viewModel.ContactIdentity == second.Identity);

        firstRequest.SetResult(new ContactDetailsDto
        {
            Identity = first.Identity,
            ContactKind = first.ContactKind,
            DisplayName = "Unexpected stale contact",
            PrimarySmtpAddress = first.PrimarySmtpAddress
        });
        await Task.Delay(100);

        Assert.Equal(second.Identity, viewModel.ContactIdentity);
        Assert.Equal(second.DisplayName, viewModel.DisplayName);
        Assert.Equal(second.PrimarySmtpAddress, viewModel.PrimarySmtpAddress);
    }

    private static async Task WaitForConditionAsync(Func<bool> condition)
    {
        for (var attempt = 0; attempt < 20; attempt++)
        {
            if (condition())
            {
                return;
            }

            await Task.Delay(25);
        }

        Assert.True(condition(), "Condition not reached in time.");
    }

    private static void SetExchangeConnected(ShellViewModel shell)
    {
        SetExchangeState(shell, ConnectionState.Connected);
    }

    private static void SetExchangeDisconnected(ShellViewModel shell)
    {
        SetExchangeState(shell, ConnectionState.Disconnected);
    }

    private static void SetExchangeState(ShellViewModel shell, ConnectionState state)
    {
        var property = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.ExchangeState), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(property);
        property!.SetValue(shell, state);
    }

    private sealed class ContactsWorkerServiceStub : TestWorkerServiceBase
    {
        public int UpsertContactCalls { get; private set; }
        public UpsertContactRequest? LastUpsertContactRequest { get; private set; }
        public List<GetContactsRequest> ContactRequests { get; } = [];
        public Queue<GetContactsResponse> ContactResponses { get; } = new();
        public List<GetContactDetailsRequest> ContactDetailRequests { get; } = [];
        public Func<GetContactDetailsRequest, Task<ContactDetailsDto>>? ContactDetailsFactory { get; set; }

        public override Task<Result> UpsertContactAsync(UpsertContactRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            UpsertContactCalls++;
            LastUpsertContactRequest = request;
            return Task.FromResult(Result.Success());
        }

        public override Task<Result<GetContactsResponse>> GetContactsAsync(GetContactsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            ContactRequests.Add(new GetContactsRequest
            {
                ContactKind = request.ContactKind,
                SearchQuery = request.SearchQuery,
                PageSize = request.PageSize,
                Skip = request.Skip,
                SortBy = request.SortBy,
                SortDescending = request.SortDescending
            });

            var response = ContactResponses.Count > 0
                ? ContactResponses.Dequeue()
                : CreateContactsResponse(request.PageSize, request.Skip, totalCount: 1, hasMore: false);

            response.PageSize = request.PageSize;
            response.Skip = request.Skip;
            return Task.FromResult(Result<GetContactsResponse>.Success(response));
        }

        public override Task<Result<ContactDetailsDto>> GetContactDetailsAsync(GetContactDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            ContactDetailRequests.Add(new GetContactDetailsRequest
            {
                Identity = request.Identity,
                ContactKind = request.ContactKind
            });

            if (ContactDetailsFactory != null)
            {
                return ContactDetailsFactory(request).ContinueWith(
                    task => Result<ContactDetailsDto>.Success(task.Result),
                    cancellationToken,
                    TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
            }

            return Task.FromResult(Result<ContactDetailsDto>.Success(new ContactDetailsDto
            {
                Identity = request.Identity,
                ContactKind = "MailUser",
                DisplayName = "Mario Rossi",
                Name = "Mario Rossi",
                Alias = "mrossi",
                PrimarySmtpAddress = "mario.rossi@contoso.com",
                ExternalEmailAddress = "mario.rossi@gmail.com",
                UserPrincipalName = "mario.rossi@contoso.com"
            }));
        }
    }

    private static GetContactsResponse CreateContactsResponse(int pageSize, int skip = 0, int totalCount = 1, bool hasMore = false)
    {
        var contacts = Enumerable.Range(skip, pageSize)
            .Select(index => new ContactListItemDto
            {
                Identity = $"contact-{index:D4}",
                ContactKind = "MailUser",
                DisplayName = $"Contact {index:D4}",
                PrimarySmtpAddress = $"contact{index:D4}@contoso.com",
                ExternalEmailAddress = $"external{index:D4}@gmail.com"
            })
            .ToList();

        return new GetContactsResponse
        {
            Contacts = contacts,
            TotalCount = totalCount,
            PageSize = pageSize,
            Skip = skip,
            HasMore = hasMore
        };
    }
}

