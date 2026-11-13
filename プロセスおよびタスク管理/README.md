# Process information scripts
<!-- 解説: この見出しは「Process information scripts」を説明する節のタイトルです。 -->

These scripts provide read-only information about processes on the local
computer. Run an individual script from Command Prompt:
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: These scripts provide read-only information about processes on the local -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: computer. Run an individual script from Command Prompt: -->

```text
cscript //nologo GetRunningProcesses.vbs
cscript //nologo TestProcessRunning.vbs notepad.exe
cscript //nologo GetProcessMemoryUsage.vbs
```
<!-- 解説: この行はREADME内の実行例コードブロックの開始を示します。 -->
<!-- 解説: このコマンド例はGetRunningProcesses.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はTestProcessRunning.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はGetProcessMemoryUsage.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: この行は実行例コードブロックの終端です。 -->

The scripts list process names and PIDs, check for a process by exact name,
and report working-set and virtual-memory sizes. They do not terminate,
modify, or remotely control processes. They do not upload information.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: The scripts list process names and PIDs, check for a process by exact name, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: and report working-set and virtual-memory sizes. They do not terminate, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: modify, or remotely control processes. They do not upload information. -->

Process termination, bulk termination, automatic restart or scheduling,
arbitrary command execution, elevation, remote control, software
installation/removal, stress testing, hidden-process inspection, and memory
dumps are intentionally not included.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Process termination, bulk termination, automatic restart or scheduling, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: arbitrary command execution, elevation, remote control, software -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: installation/removal, stress testing, hidden-process inspection, and memory -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: dumps are intentionally not included. -->

Additional read-only diagnostics include per-process CPU counters (which can
exceed 100% across cores) and process creation times. Service exit codes and
pending states can also be inspected. `ExportServiceInventory.vbs
<new-csv-path>` writes only a local CSV report and refuses to overwrite an
existing file; it does not change service configuration.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Additional read-only diagnostics include per-process CPU counters (which can -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: exceed 100% across cores) and process creation times. Service exit codes and -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: pending states can also be inspected. `ExportServiceInventory.vbs -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: <new-csv-path>` writes only a local CSV report and refuses to overwrite an -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: existing file; it does not change service configuration. -->

Other individual queries report process owners, image paths, parent PIDs,
thread/handle counts, priorities, sessions, I/O counters, CPU times, memory
peaks/private memory, service-to-process mappings, descriptions, and service
dependency references. These details can expose local usernames, paths, or
session information; use them only on systems you administer.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Other individual queries report process owners, image paths, parent PIDs, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: thread/handle counts, priorities, sessions, I/O counters, CPU times, memory -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: peaks/private memory, service-to-process mappings, descriptions, and service -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: dependency references. These details can expose local usernames, paths, or -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: session information; use them only on systems you administer. -->

## Windows service information scripts
<!-- 解説: この見出しは「Windows service information scripts」を説明する節のタイトルです。 -->

These service queries are read-only and operate on the local computer:
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: These service queries are read-only and operate on the local computer: -->

```text
cscript //nologo GetWindowsServices.vbs
cscript //nologo GetServiceStatus.vbs Spooler
cscript //nologo GetServiceStartupType.vbs Spooler
cscript //nologo CheckMiddlewareServices.vbs
cscript //nologo WaitForServiceRunning.vbs Spooler 60
```
<!-- 解説: この行はREADME内の実行例コードブロックの開始を示します。 -->
<!-- 解説: このコマンド例はGetWindowsServices.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はGetServiceStatus.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はGetServiceStartupType.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はCheckMiddlewareServices.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はWaitForServiceRunning.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: この行は実行例コードブロックの終端です。 -->

They list services, query one service by its exact service name, check common
SQL Server/IIS service states, and wait up to 300 seconds for a service to
reach `Running`. Invalid arguments, missing services, and timeouts are
reported with non-zero exit codes. They do not start, stop, reconfigure, or
remotely control services.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: They list services, query one service by its exact service name, check common -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: SQL Server/IIS service states, and wait up to 300 seconds for a service to -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: reach `Running`. Invalid arguments, missing services, and timeouts are -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: reported with non-zero exit codes. They do not start, stop, reconfigure, or -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: remotely control services. -->

Service startup/recovery/account changes, remote operations, service
installation/removal, security-service enforcement, and automatic recovery
or maintenance actions are intentionally not included.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Service startup/recovery/account changes, remote operations, service -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: installation/removal, security-service enforcement, and automatic recovery -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: or maintenance actions are intentionally not included. -->
