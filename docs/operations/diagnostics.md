# Diagnostics

Tools displays Worker supervisor, connection and prerequisite status. Its console checkbox shows/hides Worker without terminating it. The console replays retained Worker logs once on first display; its close button is disabled. Hide it through Tools.

UI, Worker and Supervisor persist timestamped JSON-line logs; paths, retention and export overrides are in [storage](../configuration/storage.md). The Logs page supports inspection/export. Preserve original diagnostic evidence when reporting a failure.

| Symptom | Action |
| --- | --- |
| Worker runs but Exchange import fails | Tools → install the exact approved module → reconnect |
| Shell configuration rejected | Check the named file/environment source in the error against [settings](../configuration/settings.md) |
| Compliance/search unavailable | Check authentication mode, RBAC and search-only session requirements in [authentication](authentication.md) |
| Operation times out or is cancelled | Check tenant state before repeating a write; cancellation does not roll back submitted changes |
| Source launch exits nonzero | Read UI/Worker/Supervisor logs; start.ps1 propagates application failure |

The repaired MainWindow language conversion uses explicit `ConverterCulture=en-US`; IT/EN still controls display culture. Startup reproduction and verification are retained in the [UI audit](../audits/2026-10-10-ui-ux-audit.md).
