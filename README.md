# ExchangeAdmin

A Windows desktop application for administering Exchange Online using PowerShell 7 and the ExchangeOnlineManagement module.

## Features

### Core Architecture
- **Worker Process Architecture**: PowerShell operations run in a separate worker process for isolation and stability
- **Interactive Authentication**: Full support for MFA and Conditional Access policies
- **Cancellation Support**: All long-running operations can be cancelled end-to-end
- **Real-time Streaming**: Progress, logs, and partial results stream in real-time
- **Automatic Recovery**: Worker process automatically restarts on crash
- **Modern WPF UI**: Clean, responsive interface

### Application Features

#### Dashboard
- Mailbox counts by type (User, Shared, Room, Equipment)
- Distribution Group counts (Static, Dynamic, Unified/M365 Groups)
- Large tenant warning (>1000 mailboxes)
- Quick navigation to lists

#### Mailboxes
- Paginated list with search (debounced)
- Filter by mailbox type
- Mailbox details view with:
  - Basic info and email addresses
  - Features: Archive, Litigation Hold, Audit, Forwarding, Quotas (with tooltips)
  - Statistics: Size, item count, last logon
  - Inbox Rules with forwarding detection
  - Auto-reply configuration
- Permissions Manager with DeltaPlan:
  - FullAccess (with AutoMapping control and toggle capability)
  - SendAs, SendOnBehalf
  - Modify AutoMapping on existing permissions
  - Preview changes before apply
  - Batch application with progress
  - Comprehensive tooltips on all controls

#### Distribution Lists
- Static Distribution Groups
- Dynamic Distribution Groups (with filter preview)
- Member management (add/remove) for static groups
- Dynamic group member preview (with performance warning)
- Paginated member list

#### Capability Detection
- Auto-detects available cmdlets after connection
- Features/columns/actions disabled based on capabilities
- Clear "Not available" tooltips for unavailable features

#### Logs Viewer
- Real-time streaming from worker process
- Filter by level (Verbose, Debug, Info, Warning, Error)
- Search filter
- Copy to clipboard
- Retry/backoff/circuit breaker state visible

## Prerequisites

### Required Software

1. **Windows 10 or later**

2. **PowerShell 7+** (required)
   - Download: https://github.com/PowerShell/PowerShell/releases
   - Or install via: `winget install Microsoft.PowerShell`

3. **ExchangeOnlineManagement Module**
   ```powershell
   # Open PowerShell 7
   Install-Module ExchangeOnlineManagement -Scope CurrentUser
   ```

4. **.NET 8 Runtime** (only for framework-dependent builds)
   - Download: https://dotnet.microsoft.com/download/dotnet/8.0
   - Self-contained builds include the runtime

### Required Permissions

- Exchange Online administrator role (or appropriate RBAC permissions)
- Ability to authenticate via browser (for MFA/Conditional Access)

## Quick Start

### Development

The fastest way to build and test:

```batch
# From project root - cleans, builds, publishes, and launches
.\build\quick-test.bat
```

**After first build:**
```batch
# Quick launch without rebuild
.\launch-app.bat
```

**Executable locations:**
- UI: `artifacts\publish\ExchangeAdmin.Presentation.exe`
- Worker: `artifacts\publish\ExchangeAdmin.Worker.exe` (auto-launched)

### Production Build

```powershell
# Framework-dependent build (~25 MB, requires .NET 8 Runtime)
.\build\build.ps1 -Clean -Publish

# Self-contained build (~205 MB, includes .NET runtime)
.\build\build.ps1 -Clean -Publish -SelfContained
```

## Running the Application

1. Launch `ExchangeAdmin.Presentation.exe`
2. Click **"Start Worker"** to start the background worker process
   - A minimized console window will appear (required for authentication)
   - Worker status should show green "Connected" within 3 seconds
3. Click **"Connect Exchange"** to authenticate
   - Browser window opens automatically for Microsoft 365 authentication
   - Complete MFA if required
   - Exchange status shows green "Connected" after successful auth
4. Navigate using the sidebar:
   - **Dashboard**: Overview and statistics
   - **Mailboxes**: User mailboxes with details
   - **Shared Mailboxes**: Shared mailbox management
   - **Distribution Lists**: Groups and membership
   - **Logs**: Real-time log viewer

### Verbose Logging

Use the **"Verbose Logging"** checkbox in the toolbar:
- **Enabled**: See all PowerShell output and detailed operation logs
- **Disabled** (default): See only important messages (Info, Warning, Error)

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Presentation (WPF)                        │
│  Shell + Navigation + ViewModels (MVVM)                     │
└────────────────────────┬────────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│            Application + Infrastructure Layers               │
│  Use Cases + IPC Client + Worker Supervisor                 │
└────────────────────────┬────────────────────────────────────┘
                         │
                    Named Pipes
                    (JSON Messages)
                         │
                         ▼
┌─────────────────────────────────────────────────────────────┐
│              Worker Process (Out-of-Process)                 │
│  PowerShell 7 Runspace + ExchangeOnlineManagement           │
│  Operation Dispatcher + Capability Detection                │
└─────────────────────────────────────────────────────────────┘
```

### Projects

| Project | Description |
|---------|-------------|
| `ExchangeAdmin.Contracts` | Versioned IPC messages, DTOs |
| `ExchangeAdmin.Domain` | Error taxonomy, retry policy, circuit breaker |
| `ExchangeAdmin.Infrastructure` | Named Pipes client, worker supervisor, heartbeat |
| `ExchangeAdmin.Application` | Use cases, worker service, capability access |
| `ExchangeAdmin.Presentation` | WPF UI with MVVM, navigation, views |
| `ExchangeAdmin.Worker` | Standalone exe: PowerShell engine, EXO commands |

### IPC Protocol

Communication between UI and Worker uses Named Pipes with JSON:

- **Request/Response**: Synchronous operations with correlation IDs
- **Event Streaming**: Real-time progress, logs, and partial output
- **Heartbeat**: Worker health monitoring (5s interval)
- **Cancellation**: Cooperative cancellation via CancellationToken

### Runtime Flow (High-Level)

1. **Presentation app starts** and initializes MVVM navigation.
2. **Worker supervisor launches** `ExchangeAdmin.Worker.exe` (visible but minimized) for OAuth browser authentication.
3. **IPC handshake completes** over named pipes; status updates in the UI.
4. **Connect Exchange** triggers PowerShell `Connect-ExchangeOnline` in the worker.
5. **Capability detection** runs after auth to enable/disable UI actions.
6. **User operations** (load mailboxes, edit permissions, etc.) stream progress/logs back to the UI.

### Key Executables

| Executable | Role | Notes |
|------------|------|-------|
| `ExchangeAdmin.Presentation.exe` | WPF UI | Launches the worker and shows status/logs |
| `ExchangeAdmin.Worker.exe` | PowerShell worker | Runs Exchange cmdlets, sends events/logs |

### Configuration & Dependencies

- **PowerShell 7**: Required for the worker runspace.
- **ExchangeOnlineManagement**: Required PowerShell module for Exchange cmdlets.
- **Execution policy**: Worker auto-sets `RemoteSigned` if needed.
- **Authentication**: OAuth browser flow requires a visible (minimized) worker window.

### Error Handling

Errors are classified and handled appropriately:

| Category | Transient | Retry | Examples |
|----------|-----------|-------|----------|
| Auth | No | No | MFA required, token expired |
| Permission | No | No | Access denied, insufficient privileges |
| Operation | No | No | Cmdlet not available, invalid parameter |
| Throttling | Yes | Yes | Rate limit, 429, retry-after |
| Transient | Yes | Yes | Timeout, network error, service unavailable |

**Retry Policy:**
- Max Retries: 3 (4 total attempts)
- Backoff: Exponential with jitter
- Max Delay: 30 seconds
- Respects server Retry-After headers

**Circuit Breaker:**
- Opens after 5 consecutive failures
- Recovery: 30 seconds before retry
- Protects against cascading failures

## Troubleshooting

### Worker Won't Start

1. Verify PowerShell 7 is installed:
   ```powershell
   pwsh --version
   ```

2. Verify ExchangeOnlineManagement module:
   ```powershell
   pwsh -Command "Get-Module -ListAvailable ExchangeOnlineManagement"
   ```

3. Check that `ExchangeAdmin.Worker.exe` exists in the same directory as the main application

### Authentication Issues

**MFA Required**
- Authentication window opens automatically in your browser
- Complete the MFA challenge
- Return to the application after successful authentication

**Conditional Access Blocked**
- Ensure you're on a compliant device
- Check with your IT administrator for Conditional Access policies

**Token Expired**
- Click "Disconnect" then "Connect" to refresh the session

### Features Not Available

If buttons or features appear disabled:
- Check the Logs tab for capability detection results
- Your Exchange role may not have the required cmdlets
- Contact your Exchange administrator for RBAC permissions

### Application Crash

If the application crashes or exits unexpectedly:
- Check the Worker console window for errors (maximize it)
- Review the Logs tab before the crash
- Ensure all prerequisites are properly installed
- Try the self-contained build if using framework-dependent

## Development

### For Developers

See **[AGENT.md](AGENT.md)** for critical implementation notes, known issues, and debugging tips.

See **[TESTING.md](TESTING.md)** for comprehensive testing procedures.

### Project Structure

```
OnlyExo365/
├── build/
│   ├── build.ps1           # Production build script
│   ├── quick-test.ps1      # Development build & test
│   └── quick-test.bat      # Batch wrapper
├── src/
│   ├── ExchangeAdmin.Contracts/
│   ├── ExchangeAdmin.Domain/
│   ├── ExchangeAdmin.Infrastructure/
│   ├── ExchangeAdmin.Application/
│   ├── ExchangeAdmin.Presentation/
│   └── ExchangeAdmin.Worker/
├── AGENT.md               # Development context for AI
├── README.md              # This file
├── TESTING.md             # Testing guide
└── ExchangeAdmin.sln      # Solution file
```

### Adding New Operations

1. Add operation type to `OperationType` enum in Contracts
2. Create request/response DTOs in Contracts
3. Add handler in `OperationDispatcher` in Worker
4. Add method to `IWorkerService` interface
5. Implement in `WorkerClient` and `WorkerService`
6. Add ViewModel commands in Presentation

### Build Scripts

- **`build\build.ps1`** - Production builds with optional self-contained mode
- **`build\quick-test.ps1`** - Development helper with clean/build/launch
- **`build\quick-test.bat`** - Batch wrapper for quick-test.ps1
- **`launch-app.bat`** - Quick launch without rebuild

## License

Proprietary - All rights reserved.
