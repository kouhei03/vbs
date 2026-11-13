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

## Additional local process examples (No. 71-170)

The numbered scripts below are separate and operate only on the local computer. Process restarts, forced termination, memory dumps, arbitrary hidden execution, GUI clicking/SendKeys/screen capture, remote execution/control, UAC bypass or credential handling, antivirus exclusions, security enforcement, persistence, self-deletion, shutdown cancellation, external HTTP, and mass deployment are not implemented. Some requested numbers therefore have no corresponding VBS file; they are listed as excluded below.

Implemented safe local examples:

| No. | Script | Behavior |
|---:|---|---|
| 71 | `071_ProcessWorkingSet.vbs` | 指定プロセスのWorking Set <!-- 解説: この行はNo.71のローカル診断スクリプトを示します。 --> |
| 72 | `072_ProcessPagefileUsage.vbs` | 指定プロセスのページファイル使用量 <!-- 解説: この行はNo.72のローカル診断スクリプトを示します。 --> |
| 73 | `073_ProcessCpuTimes.vbs` | 指定プロセスのKernel/User累計CPU時間 <!-- 解説: この行はNo.73のローカル診断スクリプトを示します。 --> |
| 74 | `074_ProcessThreadsHandles.vbs` | 指定プロセスのスレッド数とハンドル数 <!-- 解説: この行はNo.74のローカル診断スクリプトを示します。 --> |
| 75 | `075_ProcessIoBytes.vbs` | 指定プロセスのI/O読み書き量 <!-- 解説: この行はNo.75のローカル診断スクリプトを示します。 --> |
| 76 | `076_ProcessSessionId.vbs` | 指定プロセスのセッションID <!-- 解説: この行はNo.76のローカル診断スクリプトを示します。 --> |
| 77 | `077_ProcessOwner.vbs` | 指定プロセスの所有者 <!-- 解説: この行はNo.77のローカル診断スクリプトを示します。 --> |
| 78 | `078_ProcessExecutablePath.vbs` | 指定プロセスの実行ファイルパス <!-- 解説: この行はNo.78のローカル診断スクリプトを示します。 --> |
| 79 | `079_ProcessSignature.vbs` | プロセス実行ファイルのAuthenticode署名状態 <!-- 解説: この行はNo.79のローカル診断スクリプトを示します。 --> |
| 80 | `080_ProcessFileVersion.vbs` | 指定プロセスの製品/ファイルバージョン <!-- 解説: この行はNo.80のローカル診断スクリプトを示します。 --> |
| 81 | `081_ProcessCreationTime.vbs` | 指定プロセスの起動日時 <!-- 解説: この行はNo.81のローカル診断スクリプトを示します。 --> |
| 82 | `082_HighCpuMonitor.vbs` | CPU高負荷の有界サンプリング <!-- 解説: この行はNo.82のローカル診断スクリプトを示します。 --> |
| 83 | `083_MemoryTrend.vbs` | プロセスメモリの有界サンプリング <!-- 解説: この行はNo.83のローカル診断スクリプトを示します。 --> |
| 84 | `084_ProcessOpenHandleCount.vbs` | 指定プロセスのハンドル数（ハンドル名ではありません） <!-- 解説: この行はNo.84のローカル診断スクリプトを示します。 --> |
| 85 | `085_ProcessTree.vbs` | 親子プロセス関係のツリー表示 <!-- 解説: この行はNo.85のローカル診断スクリプトを示します。 --> |
| 86 | `086_WatchProcessExit.vbs` | プロセス終了の期限付き監視（再起動なし） <!-- 解説: この行はNo.86のローカル診断スクリプトを示します。 --> |
| 87 | `087_DuplicateProcessCheck.vbs` | 重複起動の検出のみ <!-- 解説: この行はNo.87のローカル診断スクリプトを示します。 --> |
| 88 | `088_ActivateExistingWindow.vbs` | 既存ウィンドウのアクティブ化 <!-- 解説: この行はNo.88のローカル診断スクリプトを示します。 --> |
| 89 | `089_TempCleanupPreview.vbs` | TEMP古いファイルの削除プレビューのみ <!-- 解説: この行はNo.89のローカル診断スクリプトを示します。 --> |
| 90 | `090_MaintenanceWindowCheck.vbs` | 保守時間帯の判定のみ <!-- 解説: この行はNo.90のローカル診断スクリプトを示します。 --> |
| 91 | `091_UpdatePreflight.vbs` | 更新前の関連プロセス確認 <!-- 解説: この行はNo.91のローカル診断スクリプトを示します。 --> |
| 92 | `092_ProcessSequencePreview.vbs` | 起動順序のプレビューのみ <!-- 解説: この行はNo.92のローカル診断スクリプトを示します。 --> |
| 95 | `095_WorkstationLockEvents.vbs` | 画面ロックイベント照会 <!-- 解説: この行はNo.95のローカル診断スクリプトを示します。 --> |
| 96 | `096_WorkstationUnlockEvents.vbs` | 画面ロック解除イベント照会 <!-- 解説: この行はNo.96のローカル診断スクリプトを示します。 --> |
| 97 | `097_ProcessConcurrencyCount.vbs` | 同時プロセス数の取得 <!-- 解説: この行はNo.97のローカル診断スクリプトを示します。 --> |
| 98 | `098_MemoryExhaustionEvents.vbs` | メモリ不足イベントの取得（ダンプなし） <!-- 解説: この行はNo.98のローカル診断スクリプトを示します。 --> |
| 99 | `099_NotRespondingProcesses.vbs` | 応答なしプロセス検出のみ <!-- 解説: この行はNo.99のローカル診断スクリプトを示します。 --> |
| 100 | `100_ProcessSummary.vbs` | プロセス集計サマリー <!-- 解説: この行はNo.100のローカル診断スクリプトを示します。 --> |
| 102 | `102_LaunchNotepadMaximized.vbs` | Notepad最大化起動の固定例 <!-- 解説: この行はNo.102のローカル診断スクリプトを示します。 --> |
| 103 | `103_LaunchNotepadMinimized.vbs` | Notepad最小化起動の固定例 <!-- 解説: この行はNo.103のローカル診断スクリプトを示します。 --> |
| 106 | `106_ActivateWindowTitle.vbs` | ウィンドウタイトルによるアクティブ化 <!-- 解説: この行はNo.106のローカル診断スクリプトを示します。 --> |
| 122 | `122_LocalTcpConnections.vbs` | ローカルTCP接続の一覧 <!-- 解説: この行はNo.122のローカル診断スクリプトを示します。 --> |
| 124 | `124_VpnAdapterStatus.vbs` | VPNアダプター状態照会のみ <!-- 解説: この行はNo.124のローカル診断スクリプトを示します。 --> |
| 126 | `126_RdpClientStatus.vbs` | ローカルRDPクライアント状態 <!-- 解説: この行はNo.126のローカル診断スクリプトを示します。 --> |
| 130 | `130_PrintSpoolerStatus.vbs` | スプーラーサービス状態照会 <!-- 解説: この行はNo.130のローカル診断スクリプトを示します。 --> |
| 135 | `135_TempProcessAudit.vbs` | TEMP配下の実行中プロセス監査 <!-- 解説: この行はNo.135のローカル診断スクリプトを示します。 --> |
| 136 | `136_AuditProcessAllowlist.vbs` | プロセス許可リスト照合のみ <!-- 解説: この行はNo.136のローカル診断スクリプトを示します。 --> |
| 138 | `138_TempProcessStartAudit.vbs` | TEMP配下プロセス起動情報照会 <!-- 解説: この行はNo.138のローカル診断スクリプトを示します。 --> |
| 143 | `143_SmartCardReaderStatus.vbs` | スマートカードリーダー状態照会 <!-- 解説: この行はNo.143のローカル診断スクリプトを示します。 --> |
| 145 | `145_SandboxFeatureStatus.vbs` | Windows Sandbox機能状態照会のみ <!-- 解説: この行はNo.145のローカル診断スクリプトを示します。 --> |
| 146 | `146_RunFixedCommand.vbs` | 固定コマンドの同期/非同期呼び出し例 <!-- 解説: この行はNo.146のローカル診断スクリプトを示します。 --> |
| 147 | `147_CaptureFixedCommandOutput.vbs` | 固定コマンドの出力/エラー取得例 <!-- 解説: この行はNo.147のローカル診断スクリプトを示します。 --> |
| 148 | `148_PowerShellDateOutput.vbs` | 固定PowerShell照会例 <!-- 解説: この行はNo.148のローカル診断スクリプトを示します。 --> |
| 151 | `151_ScheduledTaskStatus.vbs` | タスクスケジューラ状態照会のみ <!-- 解説: この行はNo.151のローカル診断スクリプトを示します。 --> |
| 152 | `152_ExitCodeBranchDemo.vbs` | 終了コード分岐デモ <!-- 解説: この行はNo.152のローカル診断スクリプトを示します。 --> |
| 153 | `153_ProcessTimingDemo.vbs` | 固定コマンドの実行時間計測 <!-- 解説: この行はNo.153のローカル診断スクリプトを示します。 --> |
| 154 | `154_SingleInstanceLock.vbs` | ロックファイルの単一起動デモ <!-- 解説: この行はNo.154のローカル診断スクリプトを示します。 --> |
| 156 | `156_ProcessEnvironmentDemo.vbs` | プロセススコープ環境変数のデモ <!-- 解説: この行はNo.156のローカル診断スクリプトを示します。 --> |
| 157 | `157_CliArgumentsDemo.vbs` | コマンドライン引数表示デモ <!-- 解説: この行はNo.157のローカル診断スクリプトを示します。 --> |
| 158 | `158_DroppedFileMetadata.vbs` | ドロップファイルのメタデータ取得 <!-- 解説: この行はNo.158のローカル診断スクリプトを示します。 --> |
| 160 | `160_ConsoleProgressDemo.vbs` | コンソール進捗表示デモ <!-- 解説: この行はNo.160のローカル診断スクリプトを示します。 --> |
| 161 | `161_LockFileStatus.vbs` | ロックファイル状態照会 <!-- 解説: この行はNo.161のローカル診断スクリプトを示します。 --> |
| 162 | `162_VerboseModeDemo.vbs` | verboseモード切替デモ <!-- 解説: この行はNo.162のローカル診断スクリプトを示します。 --> |
| 164 | `164_ShutdownEventHistory.vbs` | シャットダウン履歴照会のみ <!-- 解説: この行はNo.164のローカル診断スクリプトを示します。 --> |
| 167 | `167_AppendProcessLog.vbs` | プロセス数のローカルログ追記 <!-- 解説: この行はNo.167のローカル診断スクリプトを示します。 --> |
| 169 | `169_WaitForCpuBelowThreshold.vbs` | CPU負荷が下がるまでの期限付き待機 <!-- 解説: この行はNo.169のローカル診断スクリプトを示します。 --> |
| 170 | `170_ProcessDashboard.vbs` | プロセス一覧のローカルHTML出力 <!-- 解説: この行はNo.170のローカル診断スクリプトを示します。 --> |

Intentionally excluded numbers: 93, 94, 101, 104, 105, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 123, 125, 127, 128, 129, 131, 132, 133, 134, 137, 139, 140, 141, 142, 144, 149, 150, 155, 159, 163, 165, 166, 168. These items require process termination/restart, remote or privileged control, security-policy changes, data capture, persistence, or other operations outside this safe local scope.
