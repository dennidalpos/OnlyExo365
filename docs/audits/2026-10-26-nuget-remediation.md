# NuGet cryptography remediation — 2026-10-26

Task `SB-MIG-VERIFY-NUGET-20261026` is verified. The batch contains this P1 task only; UX-08 and CLEAN-01 remain blocked. Existing uncommitted documentation migration changes were preserved byte for byte.

## Dependency paths and change

The migration gate reported five advisories against Pkcs/Xml 10.0.12. Microsoft identifies 10.0.13 as patched:

| Package | Microsoft advisory | Reported issue |
| --- | --- | --- |
| Pkcs | [GHSA-2fmv-rv2c-59x7](https://github.com/dotnet/runtime/security/advisories/GHSA-2fmv-rv2c-59x7) | CPU exhaustion decoding attacker-controlled CMS attributes on Windows |
| Pkcs | [GHSA-846f-f764-76mx](https://github.com/dotnet/runtime/security/advisories/GHSA-846f-f764-76mx) | CPU exhaustion checking CMS signatures containing attacker-controlled X.509 data |
| Xml | [GHSA-7hqp-7c5m-4vq6](https://github.com/dotnet/runtime/security/advisories/GHSA-7hqp-7c5m-4vq6), [GHSA-c399-9853-6jmj](https://github.com/dotnet/runtime/security/advisories/GHSA-c399-9853-6jmj), [GHSA-r955-8q5j-3h83](https://github.com/dotnet/runtime/security/advisories/GHSA-r955-8q5j-3h83) | Resource exhaustion in signed XML processing and transforms |

Worker's lockfile and the [PowerShell SDK 7.6.6 manifest](https://api.nuget.org/v3-flatcontainer/microsoft.powershell.sdk/7.6.6/microsoft.powershell.sdk.nuspec) establish these minimum versions:

| Dependency path in Worker | Pkcs minimum | Xml minimum |
| --- | --- | --- |
| Microsoft.PowerShell.SDK 7.6.6 | 10.0.10 | 10.0.10 |
| System.Management.Automation 7.6.6 | 10.0.5 | — |
| Microsoft.Windows.Compatibility 10.0.5 | 10.0.5 | 10.0.5 |
| System.ServiceModel.Primitives 10.0.652802 | — | 10.0.0 |

Updated the existing direct Xml reference from 10.0.12 to 10.0.13. Its [manifest](https://api.nuget.org/v3-flatcontainer/system.security.cryptography.xml/10.0.13/system.security.cryptography.xml.nuspec) requires Pkcs 10.0.13 on net10.0, lifting both packages without another direct reference or a PowerShell SDK upgrade. NuGet regenerated Worker and Tests lockfiles; only these two package resolutions changed, including win-x64 runtime assets. Shell copies the Worker payload; Tests references Worker. Contracts' ProtectedData reference is outside these advisories and was preserved. No audit suppressions were introduced.

Production C# has no direct references to SignedCms, EnvelopedCms, SignedXml or EncryptedXml. This does not establish absence of exposure through PowerShell and imported modules; exploitability was not qualified. Remediation removes the affected package versions rather than depending on that assumption.

## Verification

Commands ran from the repository root on Windows x64. The machine has SDK 10.0.403; an isolated official SDK 10.0.401 archive was SHA-512 verified against Microsoft's release metadata. Installed runtime 10.0.14 was copied into that isolated root; the running testhost's loaded coreclr path confirmed 10.0.14. PATH/DOTNET_ROOT overrides applied only to verification processes. global.json and its disabled roll-forward policy were preserved.

| Check | Result |
| --- | --- |
| `dotnet restore OnlyExo365.sln --force-evaluate --artifacts-path artifacts/build --runtime win-x64` | Passed; lockfile changes limited to Xml/Pkcs |
| `dotnet test --no-restore --artifacts-path artifacts/build --filter 'FullyQualifiedName~PowerShellCryptographyCompatibilityTests\|FullyQualifiedName~GeneratedPowerShellExecutionTests' --nologo --disable-build-servers` | 7/7 passed |
| `pwsh -NoProfile -File build/assert-no-vulnerable-packages.ps1 -SolutionPath OnlyExo365.sln -ReportPath <evidence>/nuget-vulnerabilities.json` | Locked restore passed; no known vulnerabilities or tooling problems across four projects, including transitive packages |
| `dotnet test --nologo` | 823/823 passed; none skipped |
| `pwsh -NoProfile -File scripts/build.ps1 -Configuration Release -RuntimeIdentifier win-x64` | Passed; zero errors |
| `pwsh -NoProfile -File build/assert-architecture-constraints.ps1` | Passed |
| `git diff --check` | Passed |

Test commands additionally selected TRX output paths for retained evidence. No NU1902/NU1903 warnings occurred. Built test payload DLLs report Pkcs and Xml product version 10.0.13. The two new embedded-runspace tests exercise CMS signature verification and encryption/decryption using an ephemeral certificate without touching the certificate store, plus XML signature verification that rejects modified content.

Logs, TRX files, package audit, binary hashes, SDK/runtime evidence and the original task with acceptance results are retained in SecondBrain at `02-areas/products/OnlyExo365/history/2026-10-26-nuget-remediation/`. The isolated SDK and task scratch files remain at `C:\Users\Utente\AppData\Local\Temp\OnlyExo365-nuget-20261026`: automatic approval review rejected the bounded cleanup command with `blocked by policy` before execution. No retry was attempted; this operational residue is recorded under the existing CLEAN-01 task.

## Limits and remaining work

This qualifies offline package compatibility and the local test/build gates. It does not qualify tenant authentication, real Exchange/Graph modules, installer behavior or UX-08's DPI, accessibility and populated-grid performance matrix. Package updates do not service the separately installed .NET runtime or external PowerShell. CLEAN-01's earlier retained fixtures were not retried. The newly retained task scratch directory is excluded from the completed P1 batch; no commit, push or publication was performed.
