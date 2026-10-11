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
| Pending request capacity | 8 |
| Event count guard | Counter must remain below 10,000 |
| Pipe connection / handshake / request deadline | 10 s / 5 s / 15 min |
| Supervisor startup budget / restart budget | 90 s / 3 attempts |
| Heartbeat interval / timeout / grace | 5 s / 15 s / 5 s |
| Missed heartbeat threshold | 3; elapsed-time checks also participate in recovery |

The reader counts transport bytes while scanning each line in a 256 KiB buffer. It discards a line that exceeds the limit before allocating the full payload, then resumes at the next newline. The first-line UTF-8 BOM and CRLF are accepted; the BOM and delimiter are excluded from the JSON byte count. Outbound messages use `Encoding.UTF8.GetByteCount` before writing. This closes IPC-01; the earlier defect is recorded in the [review](../audits/2026-10-10-documentation-review.md). Console visibility uses a dedicated request after handshake, without stopping Worker.

## Requests and recovery

Lifetime begins before queue admission. Full queues return transient `ServiceUnavailable`; duplicate pending correlation IDs are rejected. Heartbeat/cancellation handling remains available while queued work waits. Cancelled or expired queued work does not dispatch.

`TimeoutMs` covers wait and execution. Deadline expiry returns `ErrorCode.Timeout` with `WasCancelled=false`; voluntary cancellation uses `WasCancelled=true`. Shell sends cancellation when its own wait expires. Neither outcome rolls back writes already submitted to Microsoft 365; writes are never retried automatically.

Server `Retry-After` is a minimum wait; only local exponential backoff is capped. Cancellation interrupts the wait. This follows [Microsoft Graph throttling guidance](https://learn.microsoft.com/en-us/graph/throttling).
