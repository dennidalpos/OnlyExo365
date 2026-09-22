# Operator Guide

Installation, initial setup, prerequisite verification, and troubleshooting for OnlyExo365 operators.

## System Requirements

- **Operating System**: Windows 10 or Windows 11 (x64)
- **Runtime**: Microsoft .NET 10 Desktop Runtime (for framework-dependent installs)
- **PowerShell**: PowerShell 7+ (`pwsh.exe` in system `PATH`)
- **Network**: Outbound HTTPS connectivity to Microsoft 365 / Entra ID endpoints
- **Identity**: Exchange Online Administrator or Global Reader credentials with appropriate RBAC roles

---

## Installation

Run `OnlyExo365.Setup.exe`:
- **Default Installation Path**: `C:\Program Files\OnlyExo365`
- **Permissions**: Requires local Administrator privileges (per-machine installation)
- **Shortcuts**: Installs Start Menu and Desktop shortcuts
- **Uninstallation**: Cleans installation binaries, logs, IPC secret tokens, and default export folders

---

## First Run & Prerequisite Bootstrap

1. Launch **OnlyExo365**.
2. Navigate to the **Tools** page.
3. Review the environment check panel for PowerShell 7, ExecutionPolicy, and module availability.
4. Bootstrap missing modules directly from PowerShell Gallery:
   - `ExchangeOnlineManagement` (`3.9.2`)
   - `Microsoft.Graph.Authentication` (`2.35.1`)
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
