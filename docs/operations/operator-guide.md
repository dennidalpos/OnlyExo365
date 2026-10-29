# Operator guide

Windows x64, outbound Microsoft 365 connectivity and operation-specific Exchange RBAC / Graph permissions are required. Use Windows 11 or a currently supported Windows 10 Enterprise/IoT LTSC edition, following [Microsoft's Exchange module prerequisites](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2). Framework-dependent packages need the .NET 10 Desktop Runtime; self-contained packages include it.

1. Run `OnlyExo365.Setup.exe` as local administrator. Default destination is `C:\Program Files\OnlyExo365`; Start Menu/Desktop shortcuts are installed.
2. Open **Tools** and check external PowerShell, ExecutionPolicy and approved modules. Install missing policy versions through Tools.
3. Configure the tenant/authentication using [settings](../configuration/settings.md), then connect. Check Shell/Worker and Exchange/Graph status separately; a running Worker does not prove a tenant connection.
4. If import fails, Tools remains available for repair. Use [authentication](authentication.md) for module/version and Compliance limits, [diagnostics](diagnostics.md) for logs/console, and [storage](../configuration/storage.md) for preserved data.

Visual workflow diagrams for all core user journeys: [Visual user guide](visual-guide.md).

## Message Trace

Use a strictly ordered interval of at most 10 days within the last 90 days, with no future dates. Today's end is the current time. The Worker accumulates records through the requested page plus one lookahead; remote V2 calls may retrieve batches of up to 5,000. Counts remain partial until exhaustion, and `HasMore` indicates another page. A stalled cursor fails explicitly.

Source: [ExoMessageTraceCommands](../../src/OnlyExo365.Worker/PowerShell/ExoMessageTraceCommands.cs); service limits: [Get-MessageTraceV2](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-messagetracev2).

## Mail flow editing

Saving a transport rule sends the editor's conditions, exceptions and actions, including empty fields and unchecked flags. Clearing a list or subject prefix removes that value; unchecking Stop rule processing or Delete message sends `false`. Properties outside the editor are omitted and preserved. The editor still requires a condition and an action. On update, state is applied through Enable/Disable-TransportRule after Set-TransportRule succeeds. If the state command fails, the earlier edit may already be saved; refresh and inspect the result before making another change. Writes are not retried automatically.

Connector editing changes the selected inbound/outbound connector without changing its Partner/OnPremises category. New connectors continue to use OnPremises; category conversion is not exposed by this editor. Accepted-domain creation/update applies Make default only when selected.

The generated commands and bound parameters have offline runspace coverage; this does not qualify tenant RBAC or the behavior of the deployed Exchange module. See [mail flow remediation](../audits/2026-10-26-mailflow-p1.md).

The rule Test action returns a result for each rule: match within the evaluated subset, no match/disabled, or indeterminate with reasons. Unresolved group membership, identities, missing condition metadata, audit mode and unsupported predicates prevent a confirmed match. This preview does not simulate Exchange delivery or rule actions; verify indeterminate rules on the tenant. Labels and the limitation notice follow the selected IT/EN language; Worker diagnostic reasons retain their original English wording.

Dashboard statistics and retention-policy caches are scoped to the current session. Changes to Worker/Exchange state or connected organization/user clear cached and displayed values; late responses from an earlier session are discarded. Connector/license read failures propagate as errors; Dashboard can instead show explicit partial-data warnings for unavailable sections. A successful empty read remains distinct from a failed read.
