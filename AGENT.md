# AGENT.md - Development Notes for Claude AI

> **Purpose**: This file contains important context, known issues, and implementation notes for future development sessions with Claude AI.
>
> **⚠️ IMPORTANT**: Do NOT generate random documentation files (*.md, changelogs, etc.) unless explicitly requested by the user. Keep documentation focused and purposeful.

## Project Overview

**ExchangeAdmin** is a WPF desktop application for managing Exchange Online (Microsoft 365) with an out-of-process PowerShell worker for executing Exchange commands.

**Architecture**:
- **Presentation** (WPF) ↔ **IPC (Named Pipes)** ↔ **Worker** (PowerShell)
- Clean Architecture: Presentation → Application → Infrastructure → Domain

## Critical Implementation Details

### 1. IPC Communication (Named Pipes)

**🔴 CRITICAL**: The IPC implementation has specific requirements that MUST be preserved:

#### Pipe Initialization Sequence
```csharp
// ❌ WRONG - Causes "Pipe is broken" error
_requestWriter = new StreamWriter(_requestPipe, Encoding.UTF8) { AutoFlush = true };

// ✅ CORRECT - Set AutoFlush AFTER creation
_requestWriter = new StreamWriter(_requestPipe, Encoding.UTF8);
// ... create all writers first ...
// Then enable AutoFlush later, OR use manual flush
```

**Why**: Setting `AutoFlush = true` in the object initializer triggers an immediate flush before the pipe is fully ready on both ends. This causes `IOException: Pipe is broken`.

#### Manual Flush Pattern
```csharp
await _requestWriter.WriteLineAsync(json).ConfigureAwait(false);
await _requestWriter.FlushAsync().ConfigureAwait(false);  // ✅ Manual flush
```

**Location**:
- `IpcServer.cs:398-399` (response writes)
- `IpcServer.cs:445-446` (event writes)
- `IpcClient.cs:285-286` (request writes)

#### Connection Synchronization
```csharp
// ✅ CRITICAL: Both pipes MUST connect before creating StreamReaders/Writers
await Task.WhenAll(
    _requestPipe.WaitForConnectionAsync(_serverCts.Token),
    _eventPipe.WaitForConnectionAsync(_serverCts.Token)
).ConfigureAwait(false);

// Only NOW create readers/writers
_requestReader = new StreamReader(_requestPipe, ...);
```

**Why**: If you create StreamWriters before both pipes are connected, the server may start writing before the client has connected its event pipe, causing race conditions.

### 2. PowerShell Worker Process

#### Execution Policy Setup
```csharp
var iss = InitialSessionState.CreateDefault();
iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.RemoteSigned;  // ✅ Required
_runspace = RunspaceFactory.CreateRunspace(iss);
```

**Why**: Default ExecutionPolicy is often `Restricted` which prevents loading the ExchangeOnlineManagement module.

**Auto-fix**: `Program.cs:11-79` includes automatic ExecutionPolicy detection and setup on Worker startup.

#### Interactive Authentication
```csharp
var startInfo = new ProcessStartInfo
{
    FileName = workerPath,
    CreateNoWindow = true,              // ✅ Hide worker console
    WindowStyle = ProcessWindowStyle.Hidden,  // ✅ Keep any window hidden
    // ...
};
```

**Why**: The worker should stay in the background while `Connect-ExchangeOnline` opens the browser for OAuth authentication.

**Location**: `WorkerSupervisor.cs:184-191`

### 3. Known Issues & Warnings

#### CA2000 Warning in App.xaml.cs
```
warning CA2000: Chiamare System.IDisposable.Dispose sull'oggetto creato da
'new LogsViewModel(_shellViewModel)' prima che tutti i relativi riferimenti
siano esterni all'ambito
```

**Status**: ⚠️ Benign, can be suppressed
**Why**: LogsViewModel is managed by the DI container lifetime and disposed with the application
**Action**: Add `#pragma warning disable CA2000` if it bothers you, or leave as-is

### 4. Verbose Logging Implementation

**Toggle**: `MainWindow.xaml:49-51` - Checkbox in toolbar
**Filter**: `ShellViewModel.cs:477-493` - Filters `LogLevel.Verbose` when disabled
**Property**: `ShellViewModel.cs:38` - `_isVerboseLoggingEnabled` field

```csharp
public void AddLog(LogLevel level, string message, string? source = null)
{
    // ✅ Filter verbose logs when disabled
    if (!_isVerboseLoggingEnabled && level == LogLevel.Verbose)
    {
        return;
    }
    // ... rest of implementation
}
```

### 5. Dashboard Warnings/Errors & Data Loading

**Warning Box**: `DashboardView.xaml:27-41` - Shows when `HasWarnings == true`
**Error Box**: `DashboardView.xaml:167-175` - Shows when `HasError == true`

```csharp
public bool HasWarnings => Warnings.Count > 0;  // ✅ Only shows if actual warnings
public bool HasError => !string.IsNullOrEmpty(ErrorMessage);  // ✅ Only shows if error
```

**Data Source**: `Stats.Warnings` from server response

### 6. WPF Value Converters

**Location**: `src/ExchangeAdmin.Presentation/Converters/BooleanConverters.cs`
**Registration**: `src/ExchangeAdmin.Presentation/Themes/DarkTheme.xaml`

Available converters:
- `BoolToVisibility` - Converts `true` → `Visible`, `false` → `Collapsed`
- `InverseBoolToVisibility` - Converts `true` → `Collapsed`, `false` → `Visible`
- `InverseBool` - Inverts boolean values
- `ListToString` - Converts `IEnumerable<string>` to comma-separated string
- `NullToVisibility` - Converts `null`/empty → `Collapsed`, non-null → `Visible`
  - Also handles empty strings and empty collections
- `ZeroToVisibility` - Converts `0` → `Collapsed`, non-zero → `Visible`
  - Supports int, long, and double types

**Usage in XAML**:
```xml
<TextBlock Visibility="{Binding HasError, Converter={StaticResource BoolToVisibility}}"/>
<GroupBox Visibility="{Binding Statistics, Converter={StaticResource NullToVisibility}}"/>
<Border Visibility="{Binding Count, Converter={StaticResource ZeroToVisibility}}"/>
```

#### ✅ FIXED: Dashboard loads after Exchange connection
**Previous Issue**: Dashboard would show "Error: Not connected to Exchange Online" at startup
**Solution**: Added PropertyChanged event listener in `App.xaml.cs:94-117` to wait for `IsExchangeConnected` before loading data

```csharp
_shellViewModel.PropertyChanged += async (s, e) =>
{
    if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected) && _shellViewModel.IsExchangeConnected)
    {
        // Reload data for the current page
        switch (_navigationService.CurrentPage)
        {
            case NavigationPage.Dashboard:
                await dashboardViewModel.LoadAsync();
                break;
            // ... other pages
        }
    }
};
```

**Result**: Dashboard now loads clean at startup, automatically loads data when Exchange connects

## Common Debugging Scenarios

### "Pipe is broken" error on startup
**Cause**: AutoFlush set in object initializer OR StreamWriters created before pipes connected
**Fix**: Review IPC initialization sequence above

### "Execution Policy" error when loading ExchangeOnlineManagement
**Cause**: Runspace created without ExecutionPolicy setting
**Fix**: Ensure `iss.ExecutionPolicy = RemoteSigned` before `CreateRunspace()`

### Connect-ExchangeOnline does nothing / times out
**Cause**: Worker process has `CreateNoWindow = true`
**Fix**: Set to `false` with `WindowStyle.Minimized`

### Authentication window opens but connection still fails
**Check**:
1. Console logs in Worker window (maximize it)
2. Look for PowerShell errors in Logs page
3. Verify ExecutionPolicy is set correctly
4. Check if ExchangeOnlineManagement module is installed

### Verbose logs always visible (toggle doesn't work)
**Check**:
1. `IsVerboseLoggingEnabled` property is bound in XAML
2. `AddLog` method filters correctly
3. UI binding is `{Binding IsVerboseLoggingEnabled}` not `{Binding IsVerboseLoggingEnabled, Mode=OneWay}`

## Build System

### Framework-Dependent Build
```powershell
.\build\build.ps1 -Publish
```
- Output: `artifacts/publish/`
- Size: ~235MB
- Requires: .NET 8 Runtime on target machine

### Self-Contained Build
```powershell
.\build\build.ps1 -Clean -Publish -SelfContained
```
- Output: `artifacts/publish/`
- Size: ~205MB
- Includes: .NET 8 Runtime (no installation needed)

### Build Script Location
`build/build.ps1` - PowerShell 7+ script with fail-fast semantics

## Testing Checklist for Future Changes

Before committing changes to IPC or Worker code:

- [ ] Test Worker startup (should complete in < 3 seconds)
- [ ] Test IPC handshake (should complete in < 2 seconds)
- [ ] Test Connect-ExchangeOnline (browser should open)
- [ ] Test Disconnect (should clean up resources)
- [ ] Test Stop Worker (should shut down gracefully)
- [ ] Test Restart Worker (should work without issues)
- [ ] Test Verbose Logging toggle (should filter in real-time)
- [ ] Check Dashboard for spurious warnings/errors
- [ ] Build Release (`.\build\build.ps1 -Publish`)
- [ ] Build Self-Contained (`.\build\build.ps1 -Clean -Publish -SelfContained`)

## Current Issues & TODOs

### Critical Issues (Must Fix)
- [x] **Missing MailboxDetailsView.xaml** - FIXED: Complete UI created with master-detail pattern
  - Solution: Created MailboxDetailsView.xaml and MailboxDetailsView.xaml.cs
  - Integrated into MailboxListView and SharedMailboxListView with 3-column grid layout
  - Features: Basic info, email addresses, features toggles, statistics, inbox rules, auto-reply, permissions manager
  - Location: `src/ExchangeAdmin.Presentation/Views/MailboxDetailsView.xaml`
  - Details panel appears on the right when a mailbox is selected, with GridSplitter for resizing

- [x] **Shared Mailbox navigation redirect** - FIXED: Now uses CurrentPage instead of hardcoded Mailboxes
  - Issue: `MailboxListViewModel.ViewDetails()` always navigated to `NavigationPage.Mailboxes`
  - Solution: Changed to use `_navigationService.CurrentPage` to support both Mailboxes and SharedMailboxes pages
  - Location: `src/ExchangeAdmin.Presentation/ViewModels/MailboxListViewModel.cs:316`

- [x] **Application crash (exit code -1073741510)** - FIXED: Synchronous cleanup in OnExit
  - Issue: `async void OnExit` caused process termination before async cleanup completed (STATUS_CONTROL_C_EXIT)
  - Solution: Changed to synchronous `OnExit` using `GetAwaiter().GetResult()` to wait for async disposal
  - Location: `src/ExchangeAdmin.Presentation/App.xaml.cs:124-148`
  - Added error handling and console logging for cleanup operations

- [x] **Worker NullReferenceException in ExoCommands** - FIXED: Added null checks before accessing PSObject.BaseObject
  - Issue: When PowerShell scripts returned null (via catch blocks), `result.Output.First().BaseObject` threw NullReferenceException
  - Solution: Added null-conditional checks in 4 locations before accessing BaseObject
  - Locations:
    - `ExoCommands.cs:507` - GetMailboxStatisticsAsync (original reported error)
    - `ExoCommands.cs:278` - GetMailboxesAsync
    - `ExoCommands.cs:390` - GetMailboxDetailsAsync
    - `ExoCommands.cs:619` - GetAutoReplyConfigurationAsync
  - Pattern: `var firstOutput = result.Output.Any() ? result.Output.First() : null;` then check `firstOutput != null`

- [x] **Details panel not visible after mailbox selection** - FIXED: Corrected DataContext binding in Visibility
  - Issue: Details panel didn't appear on right side when clicking mailboxes in both Mailboxes and Shared Mailboxes pages
  - Root Cause: After setting `DataContext="{Binding MailboxDetails}"`, the Visibility binding still used `{Binding MailboxDetails.HasDetails}` which was incorrect
  - Solution: Changed Visibility binding to `{Binding HasDetails}` (without MailboxDetails prefix) since DataContext is already set
  - Locations:
    - `MailboxListView.xaml:160` - Details panel Visibility binding
    - `SharedMailboxListView.xaml:160` - Details panel Visibility binding
  - Note: GridSplitter binding at line 156 correctly still uses `{Binding MailboxDetails.HasDetails}` because it's outside the DataContext scope

### High Priority
- [x] **Dashboard error on startup** - FIXED: Now waits for Exchange connection before loading data
  - Solution: Added PropertyChanged event listener in App.xaml.cs to reload data when IsExchangeConnected becomes true
  - No longer shows "Error: Not connected to Exchange Online" at startup

- [x] **AutoMapping modification support** - FIXED: Added ability to toggle AutoMapping on existing FullAccess permissions
  - Solution: Added `PermissionAction.Modify` enum value and implemented modify logic in Worker
  - New "Toggle AutoMap" button in FullAccess permissions DataGrid
  - Worker removes and re-adds permission with new AutoMapping value
  - Locations:
    - `ExchangeAdmin.Contracts/Dtos/MailboxDto.cs:466` - Added Modify enum
    - `ExchangeAdmin.Worker/PowerShell/ExoCommands.cs:796-802` - Modify implementation
    - `ExchangeAdmin.Presentation/ViewModels/MailboxDetailsViewModel.cs:415-441` - ModifyAutoMapping command
    - `ExchangeAdmin.Presentation/Views/MailboxDetailsView.xaml:390-401` - Toggle button UI

- [x] **ComboBox readability** - VERIFIED: ComboBoxes already have black text on white background
  - Existing style in `DarkTheme.xaml:248-357` provides optimal contrast
  - No changes needed

- [x] **Tooltips for all interactive elements** - FIXED: Added comprehensive tooltips throughout UI
  - Features section checkboxes with detailed descriptions
  - Permission manager inputs and buttons with clear action descriptions
  - Navigation buttons with functional hints
  - All buttons in DataGrids with action explanations
  - Locations:
    - `MailboxDetailsView.xaml:112,122,138,149,157` - Feature tooltips
    - `MailboxDetailsView.xaml:354,360,369,374` - Permission input tooltips
    - `MailboxDetailsView.xaml:22,33,324,326,397,407,420,432,444` - Button tooltips

- [ ] **Add unit tests** - Especially for IPC message serialization/deserialization
- [ ] **Add integration tests** - Worker startup, IPC handshake, basic Exchange commands

### Medium Priority
- [ ] **Improve error messages** - User-friendly error dialogs instead of console-only
- [ ] **Add retry logic** - Auto-retry on transient Exchange API errors
- [ ] **Performance monitoring** - Track operation durations, show in UI
- [ ] **Cache capabilities** - Don't re-detect capabilities on every connection

### Low Priority
- [ ] **Dark theme** - UI currently uses light theme only
- [ ] **Localization** - All strings are hardcoded in English/Italian
- [ ] **Keyboard shortcuts** - Add hotkeys for common actions
- [ ] **Export to CSV** - For mailbox/group lists

## File Locations Reference

### Core IPC Implementation
- **Server**: `src/ExchangeAdmin.Worker/Ipc/IpcServer.cs`
- **Client**: `src/ExchangeAdmin.Infrastructure/Ipc/IpcClient.cs`
- **Supervisor**: `src/ExchangeAdmin.Infrastructure/Ipc/WorkerSupervisor.cs`
- **Constants**: `src/ExchangeAdmin.Contracts/IpcConstants.cs`

### PowerShell Integration
- **Engine**: `src/ExchangeAdmin.Worker/PowerShell/PowerShellEngine.cs`
- **Dispatcher**: `src/ExchangeAdmin.Worker/Operations/OperationDispatcher.cs`
- **Startup**: `src/ExchangeAdmin.Worker/Program.cs`

### UI Components
- **Main Window**: `src/ExchangeAdmin.Presentation/Views/MainWindow.xaml`
- **Dashboard**: `src/ExchangeAdmin.Presentation/Views/DashboardView.xaml`
- **Shell ViewModel**: `src/ExchangeAdmin.Presentation/ViewModels/ShellViewModel.cs`

### Build & Documentation
- **Build Script**: `build/build.ps1`
- **Test Guide**: `TESTING.md`
- **Main README**: `README.md`
- **This File**: `AGENT.md`

## Architecture Decisions

### Why Out-of-Process PowerShell?
1. **Isolation**: PowerShell crashes don't bring down the UI
2. **Cancellation**: Can kill worker process if operations hang
3. **Module Loading**: Exchange module loads in separate AppDomain
4. **Security**: Credentials never in UI process memory

### Why Named Pipes?
1. **Performance**: Faster than HTTP/gRPC for local IPC
2. **Security**: No network exposure, local-only communication
3. **Reliability**: Built-in Windows mechanism with good tooling
4. **Simplicity**: No need for ports, firewall rules, or service registration

### Why Two Pipes?
1. **Request/Response**: Bidirectional pipe for commands
2. **Events**: Unidirectional pipe for async notifications (progress, logs)
3. **Prevents Deadlock**: Events can flow while request is processing
4. **Clear Semantics**: Request waits for response, events are fire-and-forget

## Security Considerations

### Current State
- ✅ No credentials stored in UI process
- ✅ Uses Windows authentication (Named Pipes)
- ✅ OAuth flow for Exchange (browser-based)
- ✅ ExecutionPolicy prevents unsigned scripts
- ✅ Worker console hidden during operation

### Future Improvements
- [ ] Add audit logging for all Exchange operations
- [ ] Implement role-based access (if multi-user support added)

## Dependencies

### Required Runtime
- **PowerShell 7.4.1+** (`pwsh.exe` in PATH)
- **ExchangeOnlineManagement 3.9.2+** (PowerShell module)
- **.NET 8.0 Runtime** (for framework-dependent builds)

### Development
- **.NET 8.0 SDK**
- **Visual Studio 2022** or **VS Code** with C# extension
- **PowerShell 7+** (for build script)

### NuGet Packages
See `Directory.Packages.props` for centralized version management

## Troubleshooting Guide for Claude

### If user reports "Worker won't start"
1. Check that `pwsh.exe` is in PATH
2. Verify ExchangeOnlineManagement module is installed
3. Look for errors in the Logs tab or captured worker output
4. Check `Program.cs` EnsureExecutionPolicyAsync logs

### If user reports "Connection times out"
1. Check IPC handshake logs
2. Verify both pipes connected (`[IPC] Client connected`)
3. Look for "Pipe is broken" errors
4. Review IPC initialization sequence in this file

### If user reports "Authentication doesn't work"
1. Check `WorkerSupervisor.cs` - ensure the worker stays hidden while launching the browser
2. Look for browser window opening (might be behind other windows)
3. Check PowerShell errors for module loading issues
4. Verify ExecutionPolicy is set correctly

### If user reports "Too many logs"
1. Check if Verbose Logging is enabled (checkbox in toolbar)
2. Suggest disabling it
3. Verify `AddLog` method filters correctly

## Notes for Future Claude Sessions

- **Don't change IPC initialization** without thorough testing
- **Don't remove manual Flush calls** - they're intentional
- **Keep the worker hidden** (`CreateNoWindow = true`, `WindowStyle = Hidden`)
- **Don't assume ExecutionPolicy is set** - verify or set it
- **Don't add AutoFlush = true** to StreamWriter initializers

When in doubt, **test the changes** with Release build before committing!

## Current Project Status

**Build Status**: ✅ Successfully compiling
**Runtime Status**: ✅ Fully functional
**Last Build**: Self-contained Release (204.99 MB) at `artifacts/publish/`

**Completed Features**:
- ✅ Dashboard with stats and warnings
- ✅ Mailbox management with master-detail UI
- ✅ Shared mailbox management
- ✅ Distribution list management
- ✅ Mailbox details panel with:
  - Basic information and email addresses
  - Feature toggles (Litigation Hold, Audit, Single Item Recovery, Retention Hold)
  - Mailbox statistics (size, item count, last logon)
  - Inbox rules with forwarding/redirect warnings
  - Auto-reply configuration
  - Permissions manager (FullAccess, SendAs, SendOnBehalf) with:
    - Add/Remove permissions with AutoMapping control
    - **NEW:** Toggle AutoMapping on existing FullAccess permissions
    - DeltaPlan for batch operations
    - Apply/Discard pending changes
- ✅ Comprehensive tooltips on all interactive elements
- ✅ Optimized ComboBox styling (black text on white background)
- ✅ Logging system with verbose toggle
- ✅ Worker process management (start/stop/restart)
- ✅ Exchange Online connection with OAuth

**Known Limitations**:
- No unit/integration tests yet
- No export to CSV functionality
- No dark theme (only light theme)
- No localization (mixed English/Italian strings)
- Worker output only visible in logs (no console window)

**Next Recommended Tasks**:
1. Add unit tests for IPC serialization/deserialization
2. Add integration tests for Worker startup and basic Exchange commands
3. Implement CSV export for mailbox and distribution list data
4. Improve error handling with user-friendly dialogs
5. Add performance monitoring and operation duration tracking

## Recent Updates (2026-01-21)

### Background Worker + Documentation
- Worker console now stays hidden during normal operation
- Documentation updated to reflect background launch behavior and log-only output

## Recent Updates (2026-01-18)

### AutoMapping Toggle Feature
- Added ability to modify AutoMapping on existing FullAccess permissions
- Implemented `PermissionAction.Modify` to handle permission updates
- Worker removes and re-adds permission with new AutoMapping value
- UI includes "Toggle AutoMap" button in permissions DataGrid
- Changes tracked in DeltaPlan with Apply/Discard workflow

### UI Improvements
- Verified ComboBox readability (black text on white background already implemented)
- Added comprehensive tooltips across all interactive elements:
  - Feature toggles explain what each mailbox feature does
  - Permission inputs describe expected values
  - Action buttons clarify their operations
- All checkboxes verified as functional with correct command bindings

### Testing
- Build completed successfully with 0 errors and 0 warnings
- All modifications tested and working correctly

---

Last Updated: 2026-01-21
By: GPT-5.2-Codex (via OpenAI)
