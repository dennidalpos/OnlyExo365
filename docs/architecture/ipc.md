# IPC

Sources: [IpcConstants](../../src/OnlyExo365.Contracts/IpcConstants.cs), [IpcSessionContext](../../src/OnlyExo365.Contracts/IpcSessionContext.cs), [IpcServer](../../src/OnlyExo365.Worker/Ipc/IpcServer.cs), [IpcClient](../../src/OnlyExo365.Shell/Ipc/IpcClient.cs), [WorkerSupervisor](../../src/OnlyExo365.Shell/Ipc/WorkerSupervisor.cs).

## Transport and security

- Windows named pipes: duplex request/response and Worker-to-Shell events. Base names `OnlyExo365_IPC_Main` / `OnlyExo365_IPC_Events` receive a user-scope SHA256 prefix and Windows session ID suffix.
- Current-user-only access; UTF-8 JSON lines, newline delimiter, compatible contract version required.
- Supervisor creates a random session token and passes it directly to the child environment as `ONLYEXO365_IPC_SESSION_TOKEN`. It is not DPAPI ciphertext. Handshake validates token, user scope and session ID; a failed handshake leaves domain requests blocked.
- DPAPI `CurrentUser` separately protects password references in `PreparedIpcPayload` / `ProtectedSecretStore`; Worker consumes and deletes those files. The session token and protected payload references are different mechanisms.

| Setting | Default |
| --- | --- |
| Message guard / pipe buffer | 10,485,760 UTF-8 bytes / 64 KiB |
| Pending request capacity / total admitted work | 8 / 9, including concurrent work |
| Event count guard | Counter must remain below 10,000 |
| Pipe connection / handshake / request deadline | 10 s / 5 s / 15 min |
| Supervisor startup budget / restart budget | 90 s / 3 attempts |
| Heartbeat interval / timeout / grace | 5 s / 15 s / 5 s |
| Missed heartbeat threshold | 3; elapsed-time checks also participate in recovery |

The reader counts transport bytes while scanning each line in a 256 KiB buffer. It discards a line that exceeds the limit before allocating the full payload, then resumes at the next newline. The first-line UTF-8 BOM and CRLF are accepted; the BOM and delimiter are excluded from the JSON byte count. Outbound messages use `Encoding.UTF8.GetByteCount` before writing. This closes IPC-01, repaired and verified in [local remediation](../audits/2026-10-11-local-remediation.md) after defect identification in the [review](../audits/2026-10-10-documentation-review.md). Console visibility uses a dedicated request after handshake, without stopping Worker.

## Requests and recovery

Lifetime begins before queue admission. Full queues return transient `ServiceUnavailable`; duplicate pending correlation IDs are rejected. Heartbeat/cancellation handling remains available while queued work waits. Cancelled or expired queued work does not dispatch.

`TimeoutMs` covers wait and execution. Deadline expiry returns `ErrorCode.Timeout` with `WasCancelled=false`; voluntary cancellation uses `WasCancelled=true`. Shell sends cancellation when its own wait expires. Neither outcome rolls back writes already submitted to Microsoft 365; writes are never retried automatically.

The client starts its local budget before waiting for the send lock, and propagates the linked token through WriteLine, Flush and response wait. Cancellation notification is best effort with its own two-second bound and does not delay completion of the expired request. Pending response/event registrations are removed for both timeout and manual cancellation. Cancellation during a client write/flush abandons the transport to prevent reuse of a possibly partial frame; cancellation while waiting for the send lock or response preserves the connection. Worker response/event sends have a five-second budget including lock wait, and abort the transport on cancellation. Controlled-stream and lock tests cover these branches; saturated real-pipe qualification remains unverified.

Runspace initialization/recreation has a thirty-second budget, and recovery shares the execution lock and caller cancellation. Replacing a runspace clears Exchange, Graph, Compliance and delegated-scope state. The external execution-policy probe has a five-second budget and terminates its owned child on expiry.

Server `Retry-After` is a minimum wait; only local exponential backoff is capped. Cancellation interrupts the wait. This follows [Microsoft Graph throttling guidance](https://learn.microsoft.com/en-us/graph/throttling).

## Transport rule updates

`UpsertTransportRuleRequest` uses nullable optional fields: omitted/null lists, strings and booleans leave existing values unchanged. An explicit empty list or empty subject prefix clears that field on update (the Worker binds `$null`); explicit `false` disables a boolean action. Name remains required. Empty optional values are omitted on creation; default mode is Enforce and default state is enabled. The Shell sends all fields it exposes, including empty values, and the Worker leaves other Exchange properties untouched. Shell and Worker ship together with this contract.

`Set-TransportRule` receives conditions/actions/mode, followed by Enable/Disable-TransportRule only when state was supplied; these writes are sequential and are not atomic. `New-TransportRule` receives Enabled directly. References: [Set-TransportRule](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-transportrule?view=exchange-ps), [New-TransportRule](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-transportrule?view=exchange-ps).
