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
