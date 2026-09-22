![OnlyExo365 app icon](src/OnlyExo365.Shell/Assets/Generated/AppIcon.png)

# OnlyExo365

OnlyExo365 is a Windows desktop application for Microsoft Exchange Online and Microsoft Graph administration, architected around an out-of-process PowerShell host.

- **`OnlyExo365.Shell`**: WPF desktop presentation, configuration loading, local SKU catalog cache, and worker supervision.
- **`OnlyExo365.Worker`**: Out-of-process PowerShell 7.6.1 execution engine.
- **`OnlyExo365.Contracts`**: Shared IPC contracts, DTOs, configuration schemas, and diagnostics.
- **`OnlyExo365.Tests`**: Automated unit, characterization, and UI regression tests.

---

## Verified Baseline

- **Platform**: Windows x64
- **Runtime & Shell**: PowerShell 7+ (`pwsh`), .NET 10 Desktop Runtime
- **Build Toolchain**: .NET SDK `10.0.401` (pinned in `global.json`, roll-forward disabled)
- **Installer**: Inno Setup 6 (required only for installer packaging)

---

## Canonical Commands

Run from the repository root:

```powershell
# Restore dependencies
pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64

# Compile Debug build
pwsh ./scripts/build.ps1 -Configuration Debug -RuntimeIdentifier win-x64

# Run shell from source
pwsh ./scripts/start.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBuild

# Clean build artifacts
pwsh ./scripts/clean.ps1

# Local verification gate (clean, build, test, architecture, security, pack)
pwsh ./scripts/gate.ps1 -RuntimeIdentifier win-x64

# Build release installer
pwsh ./scripts/pack.ps1 -Configuration Release -LockedMode -RuntimeIdentifier win-x64
```

---

## Documentation by Domain

- **Architecture**:
  - [Overview](docs/architecture/overview.md): Runtime model, project structure, and internal layering.
  - [IPC Specification](docs/architecture/ipc.md): Named pipes, protocol framing, timeouts, and session security.
- **Configuration**:
  - [Settings & Environment](docs/configuration/settings.md): Resolution hierarchy, `appsettings.json`, and `ONLYEXO365_*` variables.
- **Development**:
  - [Workflow & Tooling](docs/development/workflow.md): Local development, scripts map, and quality gates.
- **Operations**:
  - [Operator Guide](docs/operations/operator-guide.md): Installation, first-run module bootstrap, and diagnostics console.
- **Licensing**:
  - [SKU Catalog](docs/licensing/catalog.md): Embedded catalog, background caching, and refresh procedures.
- **Releases**:
  - [Packaging & Verification](docs/releases/packaging.md): Inno Setup packaging, smoke testing, and code signing.
- **Decisions**:
  - [PowerShell SDK & .NET ADR](docs/decisions/2026-04-11-powershell-sdk-major-upgrade-strategy.md): Architectural decision record.

---

## License

OnlyExo365 is distributed under the proprietary terms in [LICENSE](LICENSE).
