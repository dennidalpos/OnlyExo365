# OnlyExo365 documentation

[Italian reference](README.it.md)

Guides for Exchange Online administrators, help desks and maintainers. Start with the [product README](../README.md), [operator guide](operations/operator-guide.md) and [visual workflows](operations/visual-guide.md). Available actions depend on authentication, permissions and session capabilities.

| Area | Guide |
|---|---|
| Daily operations and confirmation | [Operator guide](operations/operator-guide.md) |
| Eight user workflows | [Visual guide](operations/visual-guide.md) |
| Approved modules, Exchange/Graph/Compliance connections | [Authentication](operations/authentication.md) |
| Worker console, persistent logs and export | [Diagnostics](operations/diagnostics.md) |
| Real public screenshot and provenance | [Assets](assets/README.md) |
| Product information copy and its limits | [Public copy](operations/locandine-copy.md) |
| Defaults, app files, machine policy and environment overrides | [Settings](configuration/settings.md) |
| Paths, DPAPI references, reports and uninstall data | [Storage](configuration/storage.md) |
| Embedded SKU data, cache and CSV parsing | [License catalog](licensing/catalog.md) |
| WPF Shell, separate Worker and Contracts | [Architecture](architecture/overview.md) |
| Named-pipe security, framing, cancellation and recovery | [IPC](architecture/ipc.md) |
| Pinned toolchain, build/test and local gates | [Development workflow](development/workflow.md) |
| Disposable installer host, test tenant and interactive UI | [Environment qualification](development/qualification.md) |
| Inno Setup and release assets | [Packaging](releases/packaging.md) |
| Authenticode signing and verification | [Signing](releases/signing.md) |

The configuration and architecture guides are maintained in English. The former Italian visual/index/public-copy pages are retained as linked counterparts. Exact source filenames and UI labels remain unchanged.

## Historical audits

The files under [audits/](audits/) describe their dated revision and scenario, not qualification of the current build. They include the October 10 critical/UI/PowerShell/documentation reviews, October 11 remediation and October 26–27 mail-flow/dependency/application reviews. Preserve their original text and evidence links.

Current behavior is established from the source revision under review. Fixture results, a disconnected screenshot, successful authentication, clean-host installation, live tenant operations and a signed release are separate forms of evidence. A Purview command response is not a general compliance certification.

Use [Contributing](../CONTRIBUTING.md) for maintenance and [Security](../SECURITY.md) for private vulnerability reporting.
