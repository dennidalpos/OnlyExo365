# Visual user workflows

[Italian reference](visual-guide.it.md)

These flows describe the implemented views, view models and Worker commands. They do not demonstrate live tenant qualification. The UI is available in English and Italian; descriptive diagram labels may differ from individual controls.

## 1. Connection

Start the Shell/Worker and inspect Tools. Approved module versions must be present; a newer version is not an automatic substitute. Review [settings](../configuration/settings.md) and [authentication](authentication.md).

~~~mermaid
flowchart LR
    Start[Start Shell and Worker] --> Policy{Connections allowed?}
    Policy -->|No| Inspect[Review machine/app policy]
    Policy -->|Yes| Modules[Check approved modules in Tools]
    Modules --> Connect[Connect using configured authentication]
    Connect --> Result{Authenticated?}
    Result -->|No or cancelled| Diagnose[Inspect diagnostic reference]
    Result -->|Yes| Capabilities[Check Exchange and Graph separately]
~~~

A running Worker does not prove Exchange access. Graph and Compliance have separate status and capability requirements.

## 2. Dashboard

Open or refresh the Dashboard after connection. It shows available mailbox/group counts, licenses and administrative users. Partial reads remain warnings; a failed read is not an empty result. The cache is session-scoped and invalidated by connection changes. This page is not a general Microsoft 365 service-health monitor.

~~~mermaid
flowchart LR
    Open[Open Dashboard] --> Load[Load available summary]
    Load --> State{Read result}
    State -->|Complete| Show[Display summary]
    State -->|Partial| Warn[Display available data and warnings]
    State -->|Failed| Error[Inspect error and retry intentionally]
    Show --> Page[Open linked detail page]
~~~

## 3. Mailboxes

Search, select a mailbox and open its details. Settings, automatic replies, delegates and licenses have separate actions; license writes additionally require Graph and its consent.

~~~mermaid
flowchart LR
    Select[Search and select] --> Edit[Edit exposed fields]
    Edit --> Save[Save that editor]
    Save --> Confirm{Review and confirm?}
    Confirm -->|Yes| Worker[Send to Worker]
    Worker --> Check[Refresh and inspect remote state]
    Confirm -->|No| Edit
~~~

Navigation with pending edits lets you remain or discard; it does not save every editor universally. A save can involve multiple operations and a later failure does not roll back earlier writes. After timeout, inspect remote state before retrying.

## 4. Message Trace

Set increasing start/end dates: at most 10 days within the last 90, without future dates. Search by supported sender/recipient/page-size filters, browse pages/status filters and inspect events for the selected trace.

~~~mermaid
flowchart LR
    Filters[Validate date and address filters] --> Search[Search page]
    Search --> Result{Result}
    Result -->|Error| Diagnose[Inspect diagnostics]
    Result -->|Empty| Empty[No messages for these filters]
    Result -->|Data| View[Inspect visible rows and events]
    View --> Next[Load next available page]
    View --> Export[Export current view to XLSX]
~~~

Counts remain partial while more pages exist. Worker lookahead determines HasMore, not the tenant-wide total. Export does not automatically download every search page.

## 5. Mail security

Supported categories include inbound/outbound spam, phishing, malware, quarantine and DKIM, subject to session capabilities. Load policies, select one and change only exposed fields.

~~~mermaid
flowchart LR
    Policies[Load and select policy] --> Edit[Edit selected fields]
    Edit --> Save[Save selected policy]
    Save --> Drafts[Review any other draft-discard warning]
    Drafts --> Confirm[Confirm intended operation]
    Confirm --> Reload[Reload and check values and warnings]
~~~

Preserve useful drafts before accepting their discard. Applying a policy is not proof of its protective effectiveness across a tenant.

## 6. Compliance

Audit searches and eDiscovery use distinct requirements. The supported eDiscovery connection is interactive delegated search authentication. Creating a search and starting it are separate actions.

~~~mermaid
flowchart LR
    Open[Open Compliance] --> Action{Task}
    Action --> Audit[Queue audit query and follow selected task]
    Action --> Search[Refresh cases searches and actions]
    Search --> Create[Create a search]
    Create --> Start[Start it separately]
    Search --> Change[Review remove purge or hold selection]
    Change --> Confirm[Confirm only intended remote operation]
~~~

Purge exposes SoftDelete and HardDelete; do not describe every request as permanent deletion. Hold requires the supported case/parameters, and listing holds may be reported unsupported. Responses do not certify legal retention or general compliance.

## 7. Mail flow

The page groups rules, connectors, accepted/remote domains, organization relationships, sharing policies, address lists, offline address books and address-book policies. Editors save their exposed fields only.

~~~mermaid
flowchart LR
    Open[Choose area and load selection] --> Action{Operation}
    Action --> Test[Test supported rule conditions]
    Test --> Match[Match no-match or indeterminate with reasons]
    Action --> Edit[Create edit delete or change state]
    Edit --> Confirm[Review summary and confirm once]
    Confirm --> Verify[Refresh and inspect remote outcome]
~~~

Emptying an exposed rule field can clear its remote value. Updating and enabling/disabling are separate commands, so partial success is possible. Test evaluates a supported subset of conditions; it does not simulate Exchange delivery or actions. See [mail-flow details](operator-guide.md#mail-flow-editing).

## 8. Tools and diagnostics

Tools remains available for prerequisite repair after tenant connection failure. Inspect modules before intentionally installing the approved version. Worker console visibility does not stop the Worker. License-catalog update reports its own status/outcome.

~~~mermaid
flowchart LR
    Tools[Open Tools] --> Modules[Inspect or repair approved modules]
    Tools --> Console[Show or hide Worker console]
    Tools --> Catalog[Update license catalog]
    Tools --> Logs[Inspect persistent logs]
    Logs --> Export[Review diagnostic archive before sharing]
~~~

Logs can contain tenant identifiers and service messages. See [diagnostics](diagnostics.md), [storage](../configuration/storage.md) and [catalog](../licensing/catalog.md).

## Source map

Connection: ShellConnectionStateViewModel/InteractiveExchangeBootstrapService; Dashboard: DashboardViewModel; mailbox editors: MailboxDetails/Permissions/LicensesViewModel; traces: MessageTraceViewModel and ExoMessageTraceCommands; security: MailSecurityViewModel; Compliance: ComplianceViewModel and its AuditQueue partial; mail flow: MailFlow/RulesViewModel and ExoTransportRuleCommands; diagnostics: Tools/LogsViewModel. All belong to the current [source tree](../../src/).
