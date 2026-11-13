# Windows event log diagnostics

The scripts in this directory perform local, read-only event-log queries.
They print event time, ID, source, and record number, not event message text,
which may contain private data. Security-log access depends on Windows
permissions.

Examples:

```text
cscript //nologo GetRecentEvents.vbs System 24
cscript //nologo FindEventsById.vbs Application 1000 72
cscript //nologo CountEventsById.vbs Application 1000 24
cscript //nologo CountFailedLogons.vbs 24
cscript //nologo GetBootShutdownHistory.vbs 168
cscript //nologo ExportEventSummary.vbs System 24 C:\Temp\system-events.csv
```

Time windows are bounded to 1-168 hours. CSV export creates a new file and
refuses to overwrite an existing one.

Additional scripts count application errors, summarize logon/logoff and
failed-logon events, query recent security/system errors, inspect event-log
size/overwrite settings, find installer errors and network-adapter changes,
and report standard crash-dump file presence.

These scripts do not write to, clear, back up, or remotely access event logs;
change audit policy; run continuously; or collect dump files. Dump inspection
only reports whether standard dump files are present.
