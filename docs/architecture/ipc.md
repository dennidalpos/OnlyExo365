# Inter-Process Communication (IPC)

OnlyExo365 isolates UI presentation (`OnlyExo365.Shell`) from PowerShell execution (`OnlyExo365.Worker`) using local Windows Named Pipes.

## Transport & Framing

- **Main Pipe (`OnlyExo365_IPC_Main`)**: Duplex request/response channel for commands, authentication, and lifecycle control.
- **Event Pipe (`OnlyExo365_IPC_Events`)**: Asynchronous worker-to-shell stream for progress notifications, log streaming, and operational warnings.
- **Framing**: UTF-8 line-delimited JSON (`\n` delimiter).
- **Buffer & Size Limits**:
  - Max message size: 10 MB (`MaxMessageSizeBytes`).
  - Pipe buffer size: 64 KB (`PipeBufferSize`).
  - Max pending requests: 8 (`MaxPendingRequests`).
  - Max events per request: 10,000 (`MaxEventsPerRequest`).

## Security Model

- **Session Token**: Authenticated via a cryptographically random token stored in DPAPI-protected user storage (`ProtectedSecretStore`).
- **Environment Variable**: Shell transmits token to worker via `ONLYEXO365_IPC_SESSION_TOKEN`.
- **Handshake**: Worker validates session token upon connection before processing any domain request. Unauthorized connections terminate immediately.

## Timeouts & Health Monitoring

- **Handshake Timeout**: 5 seconds (`HandshakeTimeoutMs`).
- **Connection Timeout**: 10 seconds (`ConnectionTimeoutMs`).
- **Request Timeout**: 15 minutes (`RequestTimeoutMs = 900000`).
- **Heartbeat Interval**: 5 seconds (`HeartbeatIntervalMs`).
- **Heartbeat Timeout**: 15 seconds (`HeartbeatTimeoutMs`).
- **Heartbeat Missed Threshold**: 3 consecutive misses (`HeartbeatMissedThreshold`) trigger worker health alert and supervisor recovery.

## Process Supervision

`WorkerSupervisor` in the shell manages the worker process:
- Spawns `OnlyExo365.Worker.exe` on startup.
- Automatically handles reconnection and restart policies (up to 3 attempts).
- Toggles console visibility via IPC handshake without terminating the background process.
