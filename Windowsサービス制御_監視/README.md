# Windows service inspection and monitoring
<!-- 解説: この見出しは「Windows service inspection and monitoring」を説明する節のタイトルです。 -->

Run each script locally with Windows Script Host. These scripts query service
configuration and status; none starts, stops, restarts, reconfigures, installs,
or removes a service.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Run each script locally with Windows Script Host. These scripts query service -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: configuration and status; none starts, stops, restarts, reconfigures, installs, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: or removes a service. -->

Examples:
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Examples: -->

```text
cscript //nologo GetAllServices.vbs
cscript //nologo GetServiceStatus.vbs Spooler
cscript //nologo WaitForServiceRunning.vbs Spooler 60
cscript //nologo WatchServiceHealth.vbs Spooler 5 60
cscript //nologo ExportServiceInventory.vbs C:\Temp\services.csv
cscript //nologo GetServiceCrashCounts.vbs 7
cscript //nologo WatchSecurityServiceStatus.vbs
```
<!-- 解説: この行はREADME内の実行例コードブロックの開始を示します。 -->
<!-- 解説: このコマンド例はGetAllServices.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はGetServiceStatus.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はWaitForServiceRunning.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はWatchServiceHealth.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はExportServiceInventory.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はGetServiceCrashCounts.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はWatchSecurityServiceStatus.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: この行は実行例コードブロックの終端です。 -->

`WatchServiceHealth.vbs` is a bounded, console-only status monitor; it does
not attempt recovery. CSV export creates a new file and refuses to overwrite
an existing one. Security-service status reports only local state.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: `WatchServiceHealth.vbs` is a bounded, console-only status monitor; it does -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: not attempt recovery. CSV export creates a new file and refuses to overwrite -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: an existing one. Security-service status reports only local state. -->

Other read-only scripts check a service's existence, startup mode, dependency
references, logon account name, exit codes, and common SQL Server/IIS states.
The crash count report counts System log events 7031 and 7034 over a 1-30 day
window without displaying event message text.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Other read-only scripts check a service's existence, startup mode, dependency -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: references, logon account name, exit codes, and common SQL Server/IIS states. -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: The crash count report counts System log events 7031 and 7034 over a 1-30 day -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: window without displaying event message text. -->

Service lifecycle changes, automatic recovery, remote control, service
installation/removal, account changes, event-log writes, and system policy
changes are intentionally excluded.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Service lifecycle changes, automatic recovery, remote control, service -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: installation/removal, account changes, event-log writes, and system policy -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: changes are intentionally excluded. -->
