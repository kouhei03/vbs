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
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: These scripts do not write to, clear, back up, or remotely access event logs; -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: change audit policy; run continuously; or collect dump files. Dump inspection -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: only reports whether standard dump files are present. -->
