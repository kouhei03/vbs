# Windows event log diagnostics
<!-- 解説: この見出しは「Windows event log diagnostics」を説明する節のタイトルです。 -->

The scripts in this directory perform local, read-only event-log queries.
They print event time, ID, source, and record number, not event message text,
which may contain private data. Security-log access depends on Windows
permissions.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: The scripts in this directory perform local, read-only event-log queries. -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: They print event time, ID, source, and record number, not event message text, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: which may contain private data. Security-log access depends on Windows -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: permissions. -->

Examples:
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Examples: -->

```text
cscript //nologo GetRecentEvents.vbs System 24
cscript //nologo FindEventsById.vbs Application 1000 72
cscript //nologo CountEventsById.vbs Application 1000 24
cscript //nologo CountFailedLogons.vbs 24
cscript //nologo GetBootShutdownHistory.vbs 168
cscript //nologo ExportEventSummary.vbs System 24 C:\Temp\system-events.csv
```
<!-- 解説: この行はREADME内の実行例コードブロックの開始を示します。 -->
<!-- 解説: このコマンド例はGetRecentEvents.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はFindEventsById.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はCountEventsById.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はCountFailedLogons.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はGetBootShutdownHistory.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はExportEventSummary.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: この行は実行例コードブロックの終端です。 -->

Time windows are bounded to 1-168 hours. CSV export creates a new file and
refuses to overwrite an existing one.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Time windows are bounded to 1-168 hours. CSV export creates a new file and -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: refuses to overwrite an existing one. -->

Additional scripts count application errors, summarize logon/logoff and
failed-logon events, query recent security/system errors, inspect event-log
size/overwrite settings, find installer errors and network-adapter changes,
and report standard crash-dump file presence.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Additional scripts count application errors, summarize logon/logoff and -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: failed-logon events, query recent security/system errors, inspect event-log -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: size/overwrite settings, find installer errors and network-adapter changes, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: and report standard crash-dump file presence. -->

These scripts do not write to, clear, back up, or remotely access event logs;
change audit policy; run continuously; or collect dump files. Dump inspection
only reports whether standard dump files are present.

## Additional examples (items 141-170)

Each numbered item has its own script. Existing general-purpose diagnostics
remain available; these numbered files provide bounded examples. Operations
that write or clear event logs, export the Security log, access remote logs,
delete or compress files, or collect dump contents are intentionally replaced
with read-only checks or previews.

| No. | Script | Behavior |
|---:|---|---|
| 141 | `141_LatestEventSummary.vbs` | Prints up to 20 event metadata rows from the last 24 hours. |
| 142 | `142_FilterEventsById.vbs` | Filters a local standard log by event ID and bounded time range. |
| 143 | `143_ApplicationErrorDetector.vbs` | Lists Application events 1000/1001 from the last 24 hours. |
| 144 | `144_CustomEventWritePreview.vbs` | Validates an event source, ID, and message length without writing. |
| 145 | `145_EventSeverityWritePreview.vbs` | Previews Information/Warning input without writing an event. |
| 146 | `146_EventLogClearPreflight.vbs` | Shows log size and retention metadata; never clears the log. |
| 147 | `147_EventLogPolicyStatus.vbs` | Lists local event log size and overwrite settings. |
| 148 | `148_CriticalEventConsoleAlert.vbs` | Prints recent System error alerts to the console only. |
| 149 | `149_ErrorCount24Hours.vbs` | Counts System and Application errors from the last 24 hours. |
| 150 | `150_RemoteLogAccessPreview.vbs` | Validates a target name; makes no network connection. |
| 151 | `151_AuditPolicyStatus.vbs` | Reads local audit policy with `auditpol /get`. |
| 152 | `152_BootShutdownTimeline.vbs` | Lists selected boot and shutdown event metadata. |
| 153 | `153_LogonLogoffTimeline.vbs` | Lists logon/logoff event IDs and times without account details. |
| 154 | `154_FailedLogonTrend.vbs` | Counts failed logons without displaying account or address data. |
| 155 | `155_SecurityLogBackupPreflight.vbs` | Checks Security log metadata; does not export or back it up. |
| 156 | `156_RunLifecycleConsoleLog.vbs` | Prints script start/end times; does not create an event source. |
| 157 | `157_MemoryDumpStatus.vbs` | Reports standard MEMORY.DMP metadata without opening the dump. |
| 158 | `158_MiniDumpInventory.vbs` | Lists minidump metadata; does not move or copy files. |
| 159 | `159_MsiErrorSummary.vbs` | Lists recent Windows Installer error metadata. |
| 160 | `160_BackupOutcomeSummary.vbs` | Lists backup-related event metadata for review. |
| 161 | `161_DiskSpaceWarningPreview.vbs` | Compares local disk free space to a threshold; console only. |
| 162 | `162_NetworkChangeTimeline.vbs` | Lists selected local network-change event metadata. |
| 163 | `163_UsbEventTimeline.vbs` | Lists Plug and Play event metadata without device identifiers. |
| 164 | `164_TaskSchedulerErrorSummary.vbs` | Lists selected Task Scheduler operational error IDs. |
| 165 | `165_SanitizedErrorSummary.vbs` | Redacts common secret fields from supplied error text. |
| 166 | `166_LogRotationPreflight.vbs` | Counts old files and bytes; does not compress or delete. |
| 167 | `167_MaskSensitiveLogText.vbs` | Masks common secret fields in supplied text. |
| 168 | `168_ServerCollectionScopePreview.vbs` | Validates a server-list file; does not contact servers. |
| 169 | `169_BoundedEventPoll.vbs` | Polls a local System event ID for at most five minutes. |
| 170 | `170_SystemDiagnosticSummary.vbs` | Creates a new local OS summary file; does not export logs or credentials. |

Security log reads may require elevated permissions and configured auditing.
The bounded poll is finite and stops automatically; it is not a resident
monitor.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: These scripts do not write to, clear, back up, or remotely access event logs; -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: change audit policy; run continuously; or collect dump files. Dump inspection -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: only reports whether standard dump files are present. -->
