# Build Scripts

This directory contains build and development scripts for ExchangeAdmin.

## Quick Reference

### Development Scripts (Root Directory)

Located in the project root for easy access:

| Script | Description | Use Case |
|--------|-------------|----------|
| `test-app.ps1` | Quick build & launch | Development testing |
| `launch-app.bat` | Launch without rebuild | Quick testing after build |

### Build Scripts (This Directory)

Located in `build/` for production builds:

| Script | Description | Use Case |
|--------|-------------|----------|
| `build.ps1` | Full production build | Release builds |
| `quick-test.ps1` | Quick build & launch | Called by test-app.ps1 |
| `quick-test.bat` | Batch launcher | Windows convenience |

## Usage Examples

### Development Workflow

```powershell
# First time or after major changes
.\test-app.ps1

# Quick testing (no rebuild)
.\launch-app.bat

# Build without cleaning (faster)
.\test-app.ps1 -SkipClean

# Just launch (skip everything else)
.\test-app.ps1 -LaunchOnly
```

### Production Builds

```powershell
# Navigate to build directory
cd build

# Framework-dependent build (requires .NET 8 Runtime)
pwsh .\build.ps1 -Configuration Release -Publish

# Self-contained build (includes .NET runtime)
pwsh .\build.ps1 -Configuration Release -Publish -SelfContained

# Clean build from scratch
pwsh .\build.ps1 -Clean -Publish
```

## Script Details

### test-app.ps1 (Root)

Main development helper script.

**Features:**
- Cleans old builds
- Compiles solution
- Publishes to artifacts/publish
- Launches application automatically
- Color-coded output
- Error handling

**Parameters:**
- `-LaunchOnly` - Skip build, just launch
- `-SkipClean` - Don't clean before building
- `-SkipBuild` - Don't build, just publish
- `-ShowPaths` - Display file paths

### launch-app.bat (Root)

Simple launcher for testing.

**Features:**
- Verifies executable exists
- Launches application
- No rebuild required

### build/quick-test.ps1

Core build logic called by test-app.ps1.

**Features:**
- Clean old artifacts
- Build solution
- Publish both projects
- Verify executables
- Launch application
- Wait for exit

**Parameters:**
- `-SkipClean` - Skip cleaning
- `-SkipBuild` - Skip building
- `-LaunchOnly` - Only launch

### build/build.ps1

Production build script with full features.

**Features:**
- Clean build support
- Self-contained builds
- Framework-dependent builds
- Version management
- Publishing
- Detailed logging

**Parameters:**
- `-Configuration` - Debug or Release (default: Release)
- `-Clean` - Clean before building
- `-Publish` - Publish after building
- `-SelfContained` - Create self-contained build

## Output Locations

All builds output to:
```
artifacts/
  publish/
    ExchangeAdmin.Presentation.exe  (Main UI)
    ExchangeAdmin.Worker.exe        (Worker process)
    [... dependencies ...]
```

## Troubleshooting

### PowerShell 7 Not Found

**Error:** `pwsh.exe` not found

**Solution:**
1. Install PowerShell 7+: https://github.com/PowerShell/PowerShell/releases
2. Or use .NET CLI directly: `dotnet build`

### Build Fails

**Error:** Build errors during compilation

**Solution:**
```powershell
# Clean everything and rebuild
cd build
pwsh .\build.ps1 -Clean -Publish
```

### Executable Not Found

**Error:** Can't find .exe after build

**Solution:**
```powershell
# Verify paths
.\test-app.ps1 -ShowPaths

# Check if publish directory exists
ls artifacts\publish\ExchangeAdmin.Presentation.exe
```

### Worker Won't Start

**Error:** Worker process crashes on launch

**Solution:**
1. Check if PowerShell 7+ is installed
2. Verify ExchangeOnlineManagement module:
   ```powershell
   pwsh -Command "Get-Module -ListAvailable ExchangeOnlineManagement"
   ```
3. Check worker console for errors

## File Permissions

If you encounter permission errors:

```powershell
# Allow script execution (run as Administrator)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Performance Tips

### Incremental Builds

For faster development iterations:

```powershell
# First build (full clean)
.\test-app.ps1

# Subsequent builds (skip clean)
.\test-app.ps1 -SkipClean

# Just launch (no rebuild)
.\launch-app.bat
```

### Parallel Builds

The build script uses `--nologo` and `--verbosity quiet` for faster, cleaner builds.

For verbose output during debugging:
```powershell
dotnet build -c Release -v detailed
```

## CI/CD Integration

For automated builds:

```yaml
# Example GitHub Actions
- name: Build and Test
  run: |
    pwsh .\build\build.ps1 -Configuration Release -Publish

# Or with .NET CLI
- name: Build
  run: dotnet build -c Release
```

## Version Management

Build versions are managed in:
- `Directory.Build.props` - Centralized version properties
- `Directory.Packages.props` - NuGet package versions

Update versions there, not in individual .csproj files.
