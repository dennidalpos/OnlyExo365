# Testing Guide

## Quick Test Procedure

### 1. Build and Run

#### Framework-Dependent Build (Requires .NET 8)
```powershell
cd <path-to-repo>\OnlyExo365
.\build\build.ps1 -Publish
cd artifacts\publish
.\ExchangeAdmin.Presentation.exe
```

#### Self-Contained Build (Standalone)
```powershell
cd <path-to-repo>\OnlyExo365
.\build\build.ps1 -Clean -Publish -SelfContained
cd artifacts\publish
.\ExchangeAdmin.Presentation.exe
```

### 2. Test Worker Connection

1. **Start Worker**
   - Click "Start Worker" button in toolbar
   - Observe minimized console window in taskbar
   - Check "Worker:" status indicator turns Green
   - Verify no errors in console

   **Expected Output:**
   ```
   [Worker] Starting ExchangeAdmin.Worker v1.0.0
   [Worker] Process ID: XXXXX
   [Worker] Checking PowerShell Execution Policy...
   [Worker] Current Execution Policy (CurrentUser): Bypass
   [Worker] Execution Policy is already acceptable
   [Worker] PowerShell initialized. Version: 7.4.1
   [Worker] Module available: True
   [IPC] Waiting for client connection...
   [IPC] Client connected
   [IPC] Readers/writers created successfully
   [IPC] Handshake completed. Compatible: True
   ```

2. **Verify IPC Connection**
   - Status should show "Worker: ● Connected"
   - No timeout errors
   - Handshake completes within 2 seconds

### 3. Test Exchange Connection

1. **Connect to Exchange Online**
   - Click "Connect Exchange" button
   - Browser window should open for Microsoft 365 authentication
   - Sign in with your M365 credentials
   - Complete MFA if required

   **Expected Behavior:**
   - Browser authentication window opens automatically
   - Console shows: `[PowerShellEngine] Connecting to Exchange Online...`
   - After successful auth: `[PowerShellEngine] Connected to Exchange Online`
   - Status changes to "Exchange: ● Connected"
   - Connected user shown in status bar

2. **Verify Connection Status**
   - Navigate to Dashboard
   - Statistics should load (mailboxes, groups, etc.)
   - No orange warnings about missing capabilities
   - No red errors in dashboard

### 4. Test Verbose Logging

1. **Enable Verbose Logging**
   - Check the "Verbose Logging" checkbox in toolbar
   - Verify message appears in Logs: "Verbose logging enabled..."

2. **Test Log Filtering**
   - Navigate to "Logs" page
   - With verbose ON: See all PowerShell verbose output
   - Uncheck "Verbose Logging"
   - Verify message: "Verbose logging disabled..."
   - Verbose messages should stop appearing

3. **Verify Log Levels**
   - Information (blue) - Important events
   - Verbose (gray) - Detailed PowerShell output (only when enabled)
   - Warning (yellow) - Non-critical issues
   - Error (red) - Failures and exceptions

### 5. Test Dashboard

1. **Load Dashboard**
   - Navigate to "Dashboard" (should be default page)
   - Click "Refresh" button
   - Verify statistics load:
     - User Mailboxes: Shows count
     - Shared Mailboxes: Shows count
     - Distribution Groups: Shows count
     - Total counts calculated correctly

2. **Verify No Spurious Warnings**
   - ❌ **NO orange warning box** should appear if connection is successful
   - ❌ **NO red error box** should appear if data loads correctly
   - ✅ Cards should show actual counts from your tenant

### 6. Test Other Features

1. **Mailboxes Page**
   - Navigate to "Mailboxes"
   - List should load with mailbox data
   - Search/filter should work

2. **Shared Mailboxes**
   - Navigate to "Shared Mailboxes"
   - Shows only shared mailboxes
   - Filtered correctly

3. **Distribution Lists**
   - Navigate to "Distribution Lists"
   - Shows distribution groups
   - Count matches dashboard

### 7. Test Disconnection and Cleanup

1. **Disconnect from Exchange**
   - Click "Disconnect" button
   - Status changes to "Exchange: ● Disconnected"
   - User info clears from status bar

2. **Stop Worker**
   - Click "Stop Worker" button
   - Console window closes
   - Status changes to "Worker: ● Stopped"
   - No errors in shutdown

3. **Restart Test**
   - Click "Start Worker" again
   - Verify clean startup
   - No leftover connections or processes

## Automated Test Checklist

- [ ] ✅ Release build completes without errors
- [ ] ✅ Self-contained build completes without errors
- [ ] ✅ Application launches successfully
- [ ] ✅ Worker starts and connects via IPC
- [ ] ✅ Handshake completes successfully
- [ ] ✅ Connect-ExchangeOnline opens browser window
- [ ] ✅ Authentication flow completes
- [ ] ✅ Exchange connection established
- [ ] ✅ Dashboard loads statistics
- [ ] ✅ No spurious warnings/errors displayed
- [ ] ✅ Verbose logging toggle works
- [ ] ✅ Log filtering applies in real-time
- [ ] ✅ All navigation pages load correctly
- [ ] ✅ Disconnect works cleanly
- [ ] ✅ Stop Worker shuts down gracefully
- [ ] ✅ Restart cycle works without issues

## Performance Benchmarks

- **Worker Startup**: < 3 seconds
- **IPC Connection**: < 2 seconds
- **Exchange Connection**: 10-15 seconds (first time), 3-5 seconds (subsequent)
- **Dashboard Load**: 2-5 seconds (depends on tenant size)
- **Mailbox List Load**: 3-10 seconds (depends on count)

## Common Issues and Solutions

### Issue: Worker fails to start
**Solution**: Check that pwsh.exe is in PATH
```powershell
Get-Command pwsh
```

### Issue: "ExchangeOnlineManagement module not available"
**Solution**: Install the module
```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

### Issue: "Execution Policy" error
**Solution**: Should auto-fix on startup. If not, run manually:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Issue: Orange warning in Dashboard
**Check**: Are you actually connected? Worker status shows Connected AND Exchange status shows Connected?

### Issue: Authentication window doesn't open
**Solution**:
1. Check Worker console for errors
2. Verify Worker window style is Minimized (not hidden)
3. Check if browser is blocking pop-ups

### Issue: "Pipe is broken" error
**Solution**: Fixed in v1.0.0 - should not occur. If it does, file a bug report with console logs.

## Regression Testing

After any code changes, run through all tests above to ensure:
1. IPC communication still works
2. PowerShell execution still works
3. Interactive authentication still works
4. UI updates correctly
5. Logging behaves as expected

## Test Environment

- **OS**: Windows 10/11 (x64)
- **PowerShell**: 7.4.1 or later
- **.NET**: 8.0 SDK (for building), 8.0 Runtime (for running framework-dependent)
- **Module**: ExchangeOnlineManagement 3.9.2 or later
- **Exchange**: Exchange Online (Microsoft 365)

## Notes for Testers

- Always test with a **real** Microsoft 365 tenant (not test/demo)
- First connection takes longer due to module loading
- Verbose logging generates A LOT of output - use sparingly
- Dashboard warnings/errors should ONLY appear when there's actual issues
- Worker console can be maximized to see detailed logs
