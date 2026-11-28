# Security Policy

## Supported versions

Only the latest commit on the `main` branch is supported.

## Reporting a vulnerability

Please do not open a public issue for security problems.

Use GitHub's private vulnerability reporting: open the **Security** tab of this repository and choose **Report a vulnerability**. Include the affected component (Shell, Worker, IPC, scripts), the authentication mode, reproduction steps and the impact you observed. Remove tenant identifiers and secrets from any evidence you attach.

You can expect an acknowledgement within a few days. Confirmed issues are fixed on `main` and disclosed after a fix is available.

## Scope notes

- OnlyExo365 acts on a Microsoft 365 tenant with the permissions of the signed-in operator; Exchange RBAC and Microsoft Graph consent decide what it can change.
- Diagnostic exports can contain tenant identifiers and service messages; review them before sharing.
- Dependencies are monitored with Dependabot and secret scanning with push protection.
