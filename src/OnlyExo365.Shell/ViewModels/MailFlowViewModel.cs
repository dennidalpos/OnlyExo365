using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;

namespace OnlyExo365.Shell.ViewModels;

public class MailFlowViewModel : ViewModelBase
{
    private const NavigationPage AlertPage = NavigationPage.MailFlow;
    private readonly ShellViewModel _shellViewModel;
    private readonly MailFlowOperationCoordinator _coordinator;
    private readonly MailFlowRulesViewModel _rules;
    private readonly MailFlowConnectorsViewModel _connectors;
    private readonly MailFlowAcceptedDomainsViewModel _acceptedDomains;
    private readonly MailFlowRemoteDomainsViewModel _remoteDomains;
    private readonly MailFlowOrganizationRelationshipsViewModel _organizationRelationships;
    private readonly MailFlowAddressListsViewModel _addressLists;
    private readonly MailFlowAddressBookPoliciesViewModel _addressBookPolicies;
    private readonly MailFlowOfflineAddressBooksViewModel _offlineAddressBooks;
    private readonly MailFlowSharingPoliciesViewModel _sharingPolicies;

    public MailFlowViewModel(IMailFlowWorkerService workerService, ShellViewModel shellViewModel)
    {
        _shellViewModel = shellViewModel;
        _coordinator = new MailFlowOperationCoordinator();

        _rules = new MailFlowRulesViewModel(workerService, shellViewModel, _coordinator, RefreshAsync);
        _connectors = new MailFlowConnectorsViewModel(workerService, shellViewModel, _coordinator, RefreshAsync);
        _acceptedDomains = new MailFlowAcceptedDomainsViewModel(workerService, shellViewModel, _coordinator, RefreshAsync);
        _remoteDomains = new MailFlowRemoteDomainsViewModel(workerService, shellViewModel, _coordinator, RefreshAsync);
        _organizationRelationships = new MailFlowOrganizationRelationshipsViewModel(workerService, shellViewModel, _coordinator, RefreshAsync);
        _addressLists = new MailFlowAddressListsViewModel(workerService, shellViewModel, _coordinator, RefreshAsync);
        _addressBookPolicies = new MailFlowAddressBookPoliciesViewModel(workerService, shellViewModel, _coordinator, RefreshAsync);
        _offlineAddressBooks = new MailFlowOfflineAddressBooksViewModel(workerService, shellViewModel, _coordinator, RefreshAsync);
        _sharingPolicies = new MailFlowSharingPoliciesViewModel(workerService, shellViewModel, _coordinator, RefreshAsync);

        _coordinator.PropertyChanged += OnChildPropertyChanged;
        _rules.PropertyChanged += OnChildPropertyChanged;
        _connectors.PropertyChanged += OnChildPropertyChanged;
        _acceptedDomains.PropertyChanged += OnChildPropertyChanged;
        _remoteDomains.PropertyChanged += OnChildPropertyChanged;
        _organizationRelationships.PropertyChanged += OnChildPropertyChanged;
        _addressLists.PropertyChanged += OnChildPropertyChanged;
        _addressBookPolicies.PropertyChanged += OnChildPropertyChanged;
        _offlineAddressBooks.PropertyChanged += OnChildPropertyChanged;
        _sharingPolicies.PropertyChanged += OnChildPropertyChanged;
        _shellViewModel.PropertyChanged += OnShellPropertyChanged;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
    }

    public bool IsLoading => _coordinator.IsLoading;
    public string? ErrorMessage => _coordinator.ErrorMessage;
    public bool HasError => _coordinator.HasError;
    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;
    public string LoadingOverlayText => _coordinator.LoadingOverlayText;

    public ObservableCollection<TransportRuleDto> TransportRules => _rules.TransportRules;
    public ObservableCollection<ConnectorDto> Connectors => _connectors.Connectors;
    public ObservableCollection<AcceptedDomainDto> AcceptedDomains => _acceptedDomains.AcceptedDomains;
    public ObservableCollection<RemoteDomainDto> RemoteDomains => _remoteDomains.RemoteDomains;
    public ObservableCollection<OrganizationRelationshipDto> OrganizationRelationships => _organizationRelationships.OrganizationRelationships;
    public ObservableCollection<AddressListDto> AddressLists => _addressLists.AddressLists;
    public ObservableCollection<AddressBookPolicyDto> AddressBookPolicies => _addressBookPolicies.AddressBookPolicies;
    public ObservableCollection<OfflineAddressBookDto> OfflineAddressBooks => _offlineAddressBooks.OfflineAddressBooks;
    public ObservableCollection<SharingPolicyDto> SharingPolicies => _sharingPolicies.SharingPolicies;

    public IReadOnlyList<string> RuleModes => _rules.RuleModes;
    public IReadOnlyList<string> ConnectorTypes => _connectors.ConnectorTypes;
    public IReadOnlyList<string> DomainTypes => _acceptedDomains.DomainTypes;
    public IReadOnlyList<string> AllowedOofTypes => _remoteDomains.AllowedOofTypes;
    public IReadOnlyList<string> FreeBusyAccessLevels => _organizationRelationships.FreeBusyAccessLevels;
    public IReadOnlyList<string> MailTipsAccessLevels => _organizationRelationships.MailTipsAccessLevels;

    public TransportRuleDto? SelectedRule { get => _rules.SelectedRule; set => _rules.SelectedRule = value; }
    public ConnectorDto? SelectedConnector { get => _connectors.SelectedConnector; set => _connectors.SelectedConnector = value; }
    public AcceptedDomainDto? SelectedDomain { get => _acceptedDomains.SelectedDomain; set => _acceptedDomains.SelectedDomain = value; }
    public RemoteDomainDto? SelectedRemoteDomain { get => _remoteDomains.SelectedRemoteDomain; set => _remoteDomains.SelectedRemoteDomain = value; }
    public OrganizationRelationshipDto? SelectedOrganizationRelationship { get => _organizationRelationships.SelectedOrganizationRelationship; set => _organizationRelationships.SelectedOrganizationRelationship = value; }
    public AddressListDto? SelectedAddressList { get => _addressLists.SelectedAddressList; set => _addressLists.SelectedAddressList = value; }
    public AddressBookPolicyDto? SelectedAddressBookPolicy { get => _addressBookPolicies.SelectedAddressBookPolicy; set => _addressBookPolicies.SelectedAddressBookPolicy = value; }
    public OfflineAddressBookDto? SelectedOfflineAddressBook { get => _offlineAddressBooks.SelectedOfflineAddressBook; set => _offlineAddressBooks.SelectedOfflineAddressBook = value; }
    public SharingPolicyDto? SelectedSharingPolicy { get => _sharingPolicies.SelectedSharingPolicy; set => _sharingPolicies.SelectedSharingPolicy = value; }

    public bool CanEditSelectedRule => _rules.CanEditSelectedRule;
    public bool CanEditSelectedConnector => _connectors.CanEditSelectedConnector;
    public bool CanEditSelectedDomain => _acceptedDomains.CanEditSelectedDomain;
    public bool CanEditSelectedRemoteDomain => _remoteDomains.CanEditSelectedRemoteDomain;
    public bool CanEditSelectedOrganizationRelationship => _organizationRelationships.CanEditSelectedOrganizationRelationship;
    public bool CanEditSelectedAddressList => _addressLists.CanEditSelectedAddressList;
    public bool CanEditSelectedAddressBookPolicy => _addressBookPolicies.CanEditSelectedAddressBookPolicy;
    public bool CanEditSelectedOfflineAddressBook => _offlineAddressBooks.CanEditSelectedOfflineAddressBook;
    public bool CanEditSelectedSharingPolicy => _sharingPolicies.CanEditSelectedSharingPolicy;
    public bool HasAddressListSectionWarning => _addressLists.HasSectionWarning;
    public string? AddressListSectionWarningMessage => _addressLists.SectionWarningMessage;
    public bool IsAddressListSectionSupported => _addressLists.IsSectionSupported;
    public bool HasOfflineAddressBookSectionWarning => _offlineAddressBooks.HasSectionWarning;
    public string? OfflineAddressBookSectionWarningMessage => _offlineAddressBooks.SectionWarningMessage;
    public bool IsOfflineAddressBookSectionSupported => _offlineAddressBooks.IsSectionSupported;

    public bool HasAnyRuleCondition => _rules.HasAnyRuleCondition;
    public bool HasAnyRuleAction => _rules.HasAnyRuleAction;
    public bool IsRuleInputValid => _rules.IsRuleInputValid;
    public bool IsRuleTestInputValid => _rules.IsRuleTestInputValid;
    public bool IsConnectorInputValid => _connectors.IsConnectorInputValid;
    public bool IsDomainInputValid => _acceptedDomains.IsDomainInputValid;
    public bool IsRemoteDomainInputValid => _remoteDomains.IsRemoteDomainInputValid;
    public bool IsOrganizationRelationshipInputValid => _organizationRelationships.IsOrganizationRelationshipInputValid;
    public bool IsAddressListInputValid => _addressLists.IsAddressListInputValid;
    public bool IsAddressBookPolicyInputValid => _addressBookPolicies.IsAddressBookPolicyInputValid;
    public bool IsOfflineAddressBookInputValid => _offlineAddressBooks.IsOfflineAddressBookInputValid;
    public bool IsSharingPolicyInputValid => _sharingPolicies.IsSharingPolicyInputValid;

    public string RuleValidationMessage => _rules.RuleValidationMessage;
    public string TestValidationMessage => _rules.TestValidationMessage;
    public string ConnectorValidationMessage => _connectors.ConnectorValidationMessage;
    public string DomainValidationMessage => _acceptedDomains.DomainValidationMessage;
    public string RemoteDomainValidationMessage => _remoteDomains.RemoteDomainValidationMessage;
    public string OrganizationRelationshipValidationMessage => _organizationRelationships.OrganizationRelationshipValidationMessage;
    public string AddressListValidationMessage => _addressLists.AddressListValidationMessage;
    public string AddressBookPolicyValidationMessage => _addressBookPolicies.AddressBookPolicyValidationMessage;
    public string OfflineAddressBookValidationMessage => _offlineAddressBooks.OfflineAddressBookValidationMessage;
    public string SharingPolicyValidationMessage => _sharingPolicies.SharingPolicyValidationMessage;

    public string? RuleIdentity { get => _rules.RuleIdentity; set => _rules.RuleIdentity = value; }
    public string RuleName { get => _rules.RuleName; set => _rules.RuleName = value; }
    public string RuleFrom { get => _rules.RuleFrom; set => _rules.RuleFrom = value; }
    public string RuleSentTo { get => _rules.RuleSentTo; set => _rules.RuleSentTo = value; }
    public string RuleSenderDomainIs { get => _rules.RuleSenderDomainIs; set => _rules.RuleSenderDomainIs = value; }
    public string RuleRecipientDomainIs { get => _rules.RuleRecipientDomainIs; set => _rules.RuleRecipientDomainIs = value; }
    public string RuleSentToMemberOf { get => _rules.RuleSentToMemberOf; set => _rules.RuleSentToMemberOf = value; }
    public string RuleSubjectContains { get => _rules.RuleSubjectContains; set => _rules.RuleSubjectContains = value; }
    public string RuleExceptIfFrom { get => _rules.RuleExceptIfFrom; set => _rules.RuleExceptIfFrom = value; }
    public string RuleExceptIfSentTo { get => _rules.RuleExceptIfSentTo; set => _rules.RuleExceptIfSentTo = value; }
    public string RuleExceptIfSenderDomainIs { get => _rules.RuleExceptIfSenderDomainIs; set => _rules.RuleExceptIfSenderDomainIs = value; }
    public string RuleExceptIfRecipientDomainIs { get => _rules.RuleExceptIfRecipientDomainIs; set => _rules.RuleExceptIfRecipientDomainIs = value; }
    public string RuleExceptIfSubjectContains { get => _rules.RuleExceptIfSubjectContains; set => _rules.RuleExceptIfSubjectContains = value; }
    public string RulePrependSubject { get => _rules.RulePrependSubject; set => _rules.RulePrependSubject = value; }
    public string RuleRedirectMessageTo { get => _rules.RuleRedirectMessageTo; set => _rules.RuleRedirectMessageTo = value; }
    public string RuleBlindCopyTo { get => _rules.RuleBlindCopyTo; set => _rules.RuleBlindCopyTo = value; }
    public string RuleAddToRecipients { get => _rules.RuleAddToRecipients; set => _rules.RuleAddToRecipients = value; }
    public bool RuleStopRuleProcessing { get => _rules.RuleStopRuleProcessing; set => _rules.RuleStopRuleProcessing = value; }
    public bool RuleDeleteMessage { get => _rules.RuleDeleteMessage; set => _rules.RuleDeleteMessage = value; }
    public string RuleMode { get => _rules.RuleMode; set => _rules.RuleMode = value; }
    public bool RuleEnabled { get => _rules.RuleEnabled; set => _rules.RuleEnabled = value; }
    public string TestSender { get => _rules.TestSender; set => _rules.TestSender = value; }
    public string TestRecipient { get => _rules.TestRecipient; set => _rules.TestRecipient = value; }
    public string TestSubject { get => _rules.TestSubject; set => _rules.TestSubject = value; }
    public string TestResult { get => _rules.TestResult; set => _rules.TestResult = value; }

    public string? ConnectorIdentity { get => _connectors.ConnectorIdentity; set => _connectors.ConnectorIdentity = value; }
    public string? ConnectorIdentityDisplay { get => _connectors.ConnectorIdentityDisplay; set => _connectors.ConnectorIdentityDisplay = value; }
    public string ConnectorType { get => _connectors.ConnectorType; set => _connectors.ConnectorType = value; }
    public string ConnectorName { get => _connectors.ConnectorName; set => _connectors.ConnectorName = value; }
    public string ConnectorComment { get => _connectors.ConnectorComment; set => _connectors.ConnectorComment = value; }
    public bool ConnectorEnabled { get => _connectors.ConnectorEnabled; set => _connectors.ConnectorEnabled = value; }
    public string ConnectorSenderDomains { get => _connectors.ConnectorSenderDomains; set => _connectors.ConnectorSenderDomains = value; }
    public string ConnectorRecipientDomains { get => _connectors.ConnectorRecipientDomains; set => _connectors.ConnectorRecipientDomains = value; }

    public string? DomainIdentity { get => _acceptedDomains.DomainIdentity; set => _acceptedDomains.DomainIdentity = value; }
    public string DomainName { get => _acceptedDomains.DomainName; set => _acceptedDomains.DomainName = value; }
    public string DomainFqdn { get => _acceptedDomains.DomainFqdn; set => _acceptedDomains.DomainFqdn = value; }
    public string DomainType { get => _acceptedDomains.DomainType; set => _acceptedDomains.DomainType = value; }
    public bool DomainMakeDefault { get => _acceptedDomains.DomainMakeDefault; set => _acceptedDomains.DomainMakeDefault = value; }

    public string? RemoteDomainIdentity { get => _remoteDomains.RemoteDomainIdentity; set => _remoteDomains.RemoteDomainIdentity = value; }
    public string RemoteDomainName { get => _remoteDomains.RemoteDomainName; set => _remoteDomains.RemoteDomainName = value; }
    public string RemoteDomainDomainName { get => _remoteDomains.RemoteDomainDomainName; set => _remoteDomains.RemoteDomainDomainName = value; }
    public string RemoteDomainAllowedOofType { get => _remoteDomains.RemoteDomainAllowedOofType; set => _remoteDomains.RemoteDomainAllowedOofType = value; }
    public bool RemoteDomainAutoReplyEnabled { get => _remoteDomains.RemoteDomainAutoReplyEnabled; set => _remoteDomains.RemoteDomainAutoReplyEnabled = value; }
    public bool RemoteDomainAutoForwardEnabled { get => _remoteDomains.RemoteDomainAutoForwardEnabled; set => _remoteDomains.RemoteDomainAutoForwardEnabled = value; }
    public bool RemoteDomainDeliveryReportEnabled { get => _remoteDomains.RemoteDomainDeliveryReportEnabled; set => _remoteDomains.RemoteDomainDeliveryReportEnabled = value; }
    public bool RemoteDomainNdrEnabled { get => _remoteDomains.RemoteDomainNdrEnabled; set => _remoteDomains.RemoteDomainNdrEnabled = value; }
    public bool RemoteDomainMeetingForwardNotificationEnabled { get => _remoteDomains.RemoteDomainMeetingForwardNotificationEnabled; set => _remoteDomains.RemoteDomainMeetingForwardNotificationEnabled = value; }
    public bool RemoteDomainTnefEnabled { get => _remoteDomains.RemoteDomainTnefEnabled; set => _remoteDomains.RemoteDomainTnefEnabled = value; }
    public bool RemoteDomainTrustedMailOutboundEnabled { get => _remoteDomains.RemoteDomainTrustedMailOutboundEnabled; set => _remoteDomains.RemoteDomainTrustedMailOutboundEnabled = value; }
    public bool RemoteDomainIsDefault { get => _remoteDomains.RemoteDomainIsDefault; set => _remoteDomains.RemoteDomainIsDefault = value; }

    public string? OrganizationRelationshipIdentity { get => _organizationRelationships.OrganizationRelationshipIdentity; set => _organizationRelationships.OrganizationRelationshipIdentity = value; }
    public string OrganizationRelationshipName { get => _organizationRelationships.OrganizationRelationshipName; set => _organizationRelationships.OrganizationRelationshipName = value; }
    public string OrganizationRelationshipDomainNames { get => _organizationRelationships.OrganizationRelationshipDomainNames; set => _organizationRelationships.OrganizationRelationshipDomainNames = value; }
    public bool OrganizationRelationshipEnabled { get => _organizationRelationships.OrganizationRelationshipEnabled; set => _organizationRelationships.OrganizationRelationshipEnabled = value; }
    public bool OrganizationRelationshipFreeBusyAccessEnabled { get => _organizationRelationships.OrganizationRelationshipFreeBusyAccessEnabled; set => _organizationRelationships.OrganizationRelationshipFreeBusyAccessEnabled = value; }
    public string OrganizationRelationshipFreeBusyAccessLevel { get => _organizationRelationships.OrganizationRelationshipFreeBusyAccessLevel; set => _organizationRelationships.OrganizationRelationshipFreeBusyAccessLevel = value; }
    public bool OrganizationRelationshipMailTipsAccessEnabled { get => _organizationRelationships.OrganizationRelationshipMailTipsAccessEnabled; set => _organizationRelationships.OrganizationRelationshipMailTipsAccessEnabled = value; }
    public string OrganizationRelationshipMailTipsAccessLevel { get => _organizationRelationships.OrganizationRelationshipMailTipsAccessLevel; set => _organizationRelationships.OrganizationRelationshipMailTipsAccessLevel = value; }
    public string OrganizationRelationshipTargetApplicationUri { get => _organizationRelationships.OrganizationRelationshipTargetApplicationUri; set => _organizationRelationships.OrganizationRelationshipTargetApplicationUri = value; }
    public string OrganizationRelationshipTargetAutodiscoverEpr { get => _organizationRelationships.OrganizationRelationshipTargetAutodiscoverEpr; set => _organizationRelationships.OrganizationRelationshipTargetAutodiscoverEpr = value; }
    public bool OrganizationRelationshipArchiveAccessEnabled { get => _organizationRelationships.OrganizationRelationshipArchiveAccessEnabled; set => _organizationRelationships.OrganizationRelationshipArchiveAccessEnabled = value; }
    public bool OrganizationRelationshipDeliveryReportEnabled { get => _organizationRelationships.OrganizationRelationshipDeliveryReportEnabled; set => _organizationRelationships.OrganizationRelationshipDeliveryReportEnabled = value; }
    public bool OrganizationRelationshipMailboxMoveEnabled { get => _organizationRelationships.OrganizationRelationshipMailboxMoveEnabled; set => _organizationRelationships.OrganizationRelationshipMailboxMoveEnabled = value; }
    public bool OrganizationRelationshipPhotosEnabled { get => _organizationRelationships.OrganizationRelationshipPhotosEnabled; set => _organizationRelationships.OrganizationRelationshipPhotosEnabled = value; }

    public string? AddressListIdentity { get => _addressLists.AddressListIdentity; set => _addressLists.AddressListIdentity = value; }
    public string AddressListName { get => _addressLists.AddressListName; set => _addressLists.AddressListName = value; }
    public string AddressListDisplayName { get => _addressLists.AddressListDisplayName; set => _addressLists.AddressListDisplayName = value; }
    public string AddressListRecipientFilter { get => _addressLists.AddressListRecipientFilter; set => _addressLists.AddressListRecipientFilter = value; }
    public string AddressListRecipientContainer { get => _addressLists.AddressListRecipientContainer; set => _addressLists.AddressListRecipientContainer = value; }
    public string AddressListIncludedRecipients { get => _addressLists.AddressListIncludedRecipients; set => _addressLists.AddressListIncludedRecipients = value; }
    public string AddressListConditionalCompany { get => _addressLists.AddressListConditionalCompany; set => _addressLists.AddressListConditionalCompany = value; }
    public string AddressListConditionalDepartment { get => _addressLists.AddressListConditionalDepartment; set => _addressLists.AddressListConditionalDepartment = value; }
    public string AddressListConditionalStateOrProvince { get => _addressLists.AddressListConditionalStateOrProvince; set => _addressLists.AddressListConditionalStateOrProvince = value; }
    public string AddressListConditionalCustomAttribute1 { get => _addressLists.AddressListConditionalCustomAttribute1; set => _addressLists.AddressListConditionalCustomAttribute1 = value; }

    public string? AddressBookPolicyIdentity { get => _addressBookPolicies.AddressBookPolicyIdentity; set => _addressBookPolicies.AddressBookPolicyIdentity = value; }
    public string AddressBookPolicyName { get => _addressBookPolicies.AddressBookPolicyName; set => _addressBookPolicies.AddressBookPolicyName = value; }
    public string AddressBookPolicyAddressLists { get => _addressBookPolicies.AddressBookPolicyAddressLists; set => _addressBookPolicies.AddressBookPolicyAddressLists = value; }
    public string AddressBookPolicyGlobalAddressList { get => _addressBookPolicies.AddressBookPolicyGlobalAddressList; set => _addressBookPolicies.AddressBookPolicyGlobalAddressList = value; }
    public string AddressBookPolicyOfflineAddressBook { get => _addressBookPolicies.AddressBookPolicyOfflineAddressBook; set => _addressBookPolicies.AddressBookPolicyOfflineAddressBook = value; }
    public string AddressBookPolicyRoomList { get => _addressBookPolicies.AddressBookPolicyRoomList; set => _addressBookPolicies.AddressBookPolicyRoomList = value; }

    public string? OfflineAddressBookIdentity { get => _offlineAddressBooks.OfflineAddressBookIdentity; set => _offlineAddressBooks.OfflineAddressBookIdentity = value; }
    public string OfflineAddressBookName { get => _offlineAddressBooks.OfflineAddressBookName; set => _offlineAddressBooks.OfflineAddressBookName = value; }
    public string OfflineAddressBookAddressLists { get => _offlineAddressBooks.OfflineAddressBookAddressLists; set => _offlineAddressBooks.OfflineAddressBookAddressLists = value; }
    public string OfflineAddressBookDiffRetentionPeriod { get => _offlineAddressBooks.OfflineAddressBookDiffRetentionPeriod; set => _offlineAddressBooks.OfflineAddressBookDiffRetentionPeriod = value; }
    public bool OfflineAddressBookIsDefault { get => _offlineAddressBooks.OfflineAddressBookIsDefault; set => _offlineAddressBooks.OfflineAddressBookIsDefault = value; }

    public string? SharingPolicyIdentity { get => _sharingPolicies.SharingPolicyIdentity; set => _sharingPolicies.SharingPolicyIdentity = value; }
    public string SharingPolicyName { get => _sharingPolicies.SharingPolicyName; set => _sharingPolicies.SharingPolicyName = value; }
    public string SharingPolicyDomains { get => _sharingPolicies.SharingPolicyDomains; set => _sharingPolicies.SharingPolicyDomains = value; }
    public bool SharingPolicyEnabled { get => _sharingPolicies.SharingPolicyEnabled; set => _sharingPolicies.SharingPolicyEnabled = value; }
    public bool SharingPolicyMakeDefault { get => _sharingPolicies.SharingPolicyMakeDefault; set => _sharingPolicies.SharingPolicyMakeDefault = value; }
    public bool SharingPolicyIsDefault { get => _sharingPolicies.SharingPolicyIsDefault; set => _sharingPolicies.SharingPolicyIsDefault = value; }

    public ICommand RefreshCommand { get; }
    public ICommand NewRuleCommand => _rules.NewRuleCommand;
    public ICommand EnableRuleCommand => _rules.EnableRuleCommand;
    public ICommand DisableRuleCommand => _rules.DisableRuleCommand;
    public ICommand SaveRuleCommand => _rules.SaveRuleCommand;
    public ICommand RemoveRuleCommand => _rules.RemoveRuleCommand;
    public ICommand TestRuleCommand => _rules.TestRuleCommand;
    public ICommand NewConnectorCommand => _connectors.NewConnectorCommand;
    public ICommand SaveConnectorCommand => _connectors.SaveConnectorCommand;
    public ICommand RemoveConnectorCommand => _connectors.RemoveConnectorCommand;
    public ICommand NewDomainCommand => _acceptedDomains.NewDomainCommand;
    public ICommand SaveDomainCommand => _acceptedDomains.SaveDomainCommand;
    public ICommand RemoveDomainCommand => _acceptedDomains.RemoveDomainCommand;
    public ICommand NewRemoteDomainCommand => _remoteDomains.NewRemoteDomainCommand;
    public ICommand SaveRemoteDomainCommand => _remoteDomains.SaveRemoteDomainCommand;
    public ICommand RemoveRemoteDomainCommand => _remoteDomains.RemoveRemoteDomainCommand;
    public ICommand NewOrganizationRelationshipCommand => _organizationRelationships.NewOrganizationRelationshipCommand;
    public ICommand SaveOrganizationRelationshipCommand => _organizationRelationships.SaveOrganizationRelationshipCommand;
    public ICommand RemoveOrganizationRelationshipCommand => _organizationRelationships.RemoveOrganizationRelationshipCommand;
    public ICommand NewAddressListCommand => _addressLists.NewAddressListCommand;
    public ICommand SaveAddressListCommand => _addressLists.SaveAddressListCommand;
    public ICommand RemoveAddressListCommand => _addressLists.RemoveAddressListCommand;
    public ICommand NewAddressBookPolicyCommand => _addressBookPolicies.NewAddressBookPolicyCommand;
    public ICommand SaveAddressBookPolicyCommand => _addressBookPolicies.SaveAddressBookPolicyCommand;
    public ICommand RemoveAddressBookPolicyCommand => _addressBookPolicies.RemoveAddressBookPolicyCommand;
    public ICommand NewOfflineAddressBookCommand => _offlineAddressBooks.NewOfflineAddressBookCommand;
    public ICommand SaveOfflineAddressBookCommand => _offlineAddressBooks.SaveOfflineAddressBookCommand;
    public ICommand RemoveOfflineAddressBookCommand => _offlineAddressBooks.RemoveOfflineAddressBookCommand;
    public ICommand NewSharingPolicyCommand => _sharingPolicies.NewSharingPolicyCommand;
    public ICommand SaveSharingPolicyCommand => _sharingPolicies.SaveSharingPolicyCommand;
    public ICommand RemoveSharingPolicyCommand => _sharingPolicies.RemoveSharingPolicyCommand;

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            _coordinator.ClearError();
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        if (TransportRules.Count == 0 &&
            Connectors.Count == 0 &&
            AcceptedDomains.Count == 0 &&
            RemoteDomains.Count == 0 &&
            OrganizationRelationships.Count == 0 &&
            AddressLists.Count == 0 &&
            AddressBookPolicies.Count == 0 &&
            OfflineAddressBooks.Count == 0 &&
            SharingPolicies.Count == 0)
        {
            await RefreshAsync(CancellationToken.None);
        }
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        var hadWorkspaceData = HasWorkspaceData;
        _coordinator.BeginOperation("Loading Mail Flow workspace...");
        _coordinator.ClearError();
        _shellViewModel.ClearPageAlert(AlertPage);

        try
        {
            await Task.WhenAll(
                _rules.LoadAsync(cancellationToken),
                _connectors.LoadAsync(cancellationToken),
                _acceptedDomains.LoadAsync(cancellationToken),
                _remoteDomains.LoadAsync(cancellationToken),
                _organizationRelationships.LoadAsync(cancellationToken),
                _addressLists.LoadAsync(cancellationToken),
                _addressBookPolicies.LoadAsync(cancellationToken),
                _offlineAddressBooks.LoadAsync(cancellationToken),
                _sharingPolicies.LoadAsync(cancellationToken));

            if (!HasError)
            {
                _shellViewModel.ClearPageAlert(AlertPage);
                _shellViewModel.AddLog(
                    LogLevel.Information,
                    $"MailFlow refresh complete: rules={TransportRules.Count}, connectors={Connectors.Count}, domains={AcceptedDomains.Count}, remoteDomains={RemoteDomains.Count}, organizationRelationships={OrganizationRelationships.Count}, addressLists={AddressLists.Count}, addressBookPolicies={AddressBookPolicies.Count}, offlineAddressBooks={OfflineAddressBooks.Count}, sharingPolicies={SharingPolicies.Count}",
                    "MailFlow");
            }
            else if (!hadWorkspaceData && !HasWorkspaceData)
            {
                var errorMessage = ErrorMessage ?? "Unable to load the Mail Flow workspace.";
                _coordinator.ClearError();
                _shellViewModel.ShowPageLoadFailedAlert(AlertPage, errorMessage);
            }
        }
        catch (Exception ex)
        {
            if (hadWorkspaceData || HasWorkspaceData)
            {
                _coordinator.SetError(ex.Message);
            }
            else
            {
                _coordinator.ClearError();
                _shellViewModel.ShowPageLoadFailedAlert(AlertPage, ex.Message);
            }

            _shellViewModel.AddLog(LogLevel.Error, $"MailFlow refresh exception: {ex.Message}", "MailFlow");
        }
        finally
        {
            _coordinator.EndOperation();
        }
    }

    private void OnShellPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected))
        {
            if (!_shellViewModel.IsExchangeConnected)
            {
                _coordinator.ClearError();
                _shellViewModel.ClearPageAlert(AlertPage);
            }
            OnPropertyChanged(nameof(CanRefresh));
            CommandManager.InvalidateRequerySuggested();
        }
    }

    private void OnChildPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(e.PropertyName))
        {
            OnPropertyChanged(e.PropertyName);

            if (e.PropertyName == nameof(IsLoading))
            {
                OnPropertyChanged(nameof(CanRefresh));
                OnPropertyChanged(nameof(CanEditSelectedRule));
                OnPropertyChanged(nameof(CanEditSelectedConnector));
                OnPropertyChanged(nameof(CanEditSelectedDomain));
                OnPropertyChanged(nameof(CanEditSelectedRemoteDomain));
                OnPropertyChanged(nameof(CanEditSelectedOrganizationRelationship));
                OnPropertyChanged(nameof(CanEditSelectedAddressList));
                OnPropertyChanged(nameof(CanEditSelectedAddressBookPolicy));
                OnPropertyChanged(nameof(CanEditSelectedOfflineAddressBook));
                OnPropertyChanged(nameof(CanEditSelectedSharingPolicy));
            }
        }

        if (ReferenceEquals(sender, _addressLists) &&
            (e.PropertyName == nameof(MailFlowSectionViewModelBase.HasSectionWarning) ||
             e.PropertyName == nameof(MailFlowSectionViewModelBase.SectionWarningMessage) ||
             e.PropertyName == nameof(MailFlowSectionViewModelBase.IsSectionSupported)))
        {
            OnPropertyChanged(nameof(HasAddressListSectionWarning));
            OnPropertyChanged(nameof(AddressListSectionWarningMessage));
            OnPropertyChanged(nameof(IsAddressListSectionSupported));
            OnPropertyChanged(nameof(CanEditSelectedAddressList));
        }

        if (ReferenceEquals(sender, _offlineAddressBooks) &&
            (e.PropertyName == nameof(MailFlowSectionViewModelBase.HasSectionWarning) ||
             e.PropertyName == nameof(MailFlowSectionViewModelBase.SectionWarningMessage) ||
             e.PropertyName == nameof(MailFlowSectionViewModelBase.IsSectionSupported)))
        {
            OnPropertyChanged(nameof(HasOfflineAddressBookSectionWarning));
            OnPropertyChanged(nameof(OfflineAddressBookSectionWarningMessage));
            OnPropertyChanged(nameof(IsOfflineAddressBookSectionSupported));
            OnPropertyChanged(nameof(CanEditSelectedOfflineAddressBook));
        }
    }

    private bool HasWorkspaceData =>
        TransportRules.Count > 0 ||
        Connectors.Count > 0 ||
        AcceptedDomains.Count > 0 ||
        RemoteDomains.Count > 0 ||
        OrganizationRelationships.Count > 0 ||
        AddressLists.Count > 0 ||
        AddressBookPolicies.Count > 0 ||
        OfflineAddressBooks.Count > 0 ||
        SharingPolicies.Count > 0;
}


