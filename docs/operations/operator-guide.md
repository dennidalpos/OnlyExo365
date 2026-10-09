# Operator Guide

Installation, initial setup, prerequisite verification, and troubleshooting for OnlyExo365 operators.

## System Requirements

- **Operating System**: Windows 11 (x64); Windows 10 only on supported Enterprise/IoT LTSC editions
- **Runtime**: Microsoft .NET 10 Desktop Runtime (for framework-dependent installs)
- **PowerShell**: External PowerShell 7.6+ (`pwsh.exe` in system `PATH`) for interactive Exchange bootstrap and tenant validation. The Worker embeds PowerShell 7.6.6.
- **Network**: Outbound HTTPS connectivity to Microsoft 365 / Entra ID endpoints
- **Identity**: Exchange Online Administrator or Global Reader credentials with appropriate RBAC roles

---

## Installation

Run `OnlyExo365.Setup.exe`:
- **Default Installation Path**: `C:\Program Files\OnlyExo365`
- **Permissions**: Requires local Administrator privileges (per-machine installation)
- **Shortcuts**: Installs Start Menu and Desktop shortcuts
- **Uninstallation**: Removes installed binaries, logs, and IPC secret tokens; preserves exports and files added by the operator

---

## First Run & Prerequisite Bootstrap

1. Launch **OnlyExo365**.
2. Navigate to the **Tools** page.
3. Review the environment check panel for PowerShell 7, ExecutionPolicy, and module availability.
4. Bootstrap missing modules directly from PowerShell Gallery:
   - `ExchangeOnlineManagement` (`3.10.1`)
   - `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Users.Actions`, and `Microsoft.Graph.Identity.DirectoryManagement` (all `2.41.1`)
   - Graph submodules: `Microsoft.Graph.Users`, `Microsoft.Graph.Users.Actions`, `Microsoft.Graph.Identity.DirectoryManagement`
5. Connect to Exchange Online:
   - Provide tenant credentials / certificate details as configured.
   - Verify that Shell, Worker, and Exchange statuses display green/connected.

---

## Worker Diagnostics & Console

The **Tools** page provides real-time supervisor and worker status:

- **Worker Console Toggle**: Use the checkbox on the Tools page to show or hide the worker console window.
- **Log Replay**: Upon opening, the console automatically replays retained logs from `%LocalAppData%\OnlyExo365\logs\worker-*.log`.
- **Termination Protection**: The worker console window's close (`[X]`) button is programmatically disabled to prevent unintended process termination. Always use the UI checkbox to hide the console.
- **Persistent Logs**: Structured, timestamped logs for UI, Worker, and Supervisor are stored under `%LocalAppData%\OnlyExo365\logs\`.

## Approved Modules and Uninstall

Tools installs the approved versions side by side. A newer installed version does not substitute for the policy version. Missing approved modules or failed imports leave the Worker and Tools available, but block the affected connection with an explicit error. Reinstall the approved version from Tools and reconnect.

Uninstall removes installed files and transient logs/IPC secrets. Exports under `%LocalAppData%\OnlyExo365\exports` and files added by the operator inside the installation directory are preserved. Back up reports before machine maintenance.

## Message Trace

Choose a strictly ordered interval of at most 10 days, within the last 90 days, without future dates. Today's upper bound is the current time. Each search loads only enough records for the requested page plus one lookahead record. Counts remain partial until the end of the result set is reached; `HasMore` indicates whether another page exists.

Reference: [Exchange Online PowerShell prerequisites](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2), [Get-MessageTraceV2 limits](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-messagetracev2).
