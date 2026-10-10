# Signing

The [signed-release workflow](../../.github/workflows/release-signed.yml) runs the shared baseline, signs publish executables, rebuilds setup from that payload, signs setup, verifies both, then runs isolated smoke and prepares/uploads release assets. Signing only an existing setup cannot update embedded application binaries.

| Script | Purpose |
| --- | --- |
| `build/resolve-signtool.ps1` | Resolve Windows SDK SignTool |
| `build/sign-artifacts.ps1` | Authenticode signing |
| `build/create-setup-exe.ps1` | Rebuild installer from the signed publish tree |
| `build/verify-signatures.ps1` | Validate signatures and optional expected thumbprint |
| `build/validate-artifact-signing.ps1` | Sign/timestamp/verify disposable copies with a temporary certificate |

Disposable validation preserves source hashes and deletes only its own certificate thumbprint/private key. It adds no trust roots and does not establish trusted release signatures. Existing same-subject certificates survive; cleanup failure is reported.

Isolated reinstall must compare Shell/Worker SHA256 with publish and require valid installed EXE signatures matching signed setup. The generated Inno uninstaller is excluded. Unsigned smoke packages report `installed_signatures_confirmed=false`. Trusted installed-payload qualification remains AUD-07; see [qualification](../development/qualification.md).

References: [Inno Setup SignTool](https://jrsoftware.org/ishelp/topic_setup_signtool.htm), [Microsoft SignTool](https://learn.microsoft.com/en-us/windows/win32/seccrypto/signtool).
