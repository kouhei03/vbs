> **ファイル概要:** プロセスとサービスの診断スクリプトの一覧と制限事項。
> **実行方法:** この文書を参照し、Windowsのコマンドプロンプトでこのフォルダーから `cscript //nologo GetRunningProcesses.vbs` を実行。
> **実行結果:** 例では実行中のプロセス名とIDがコンソールに表示されます。

# Process information scripts

These scripts provide read-only information about processes on the local
computer. Run an individual script from Command Prompt:

```text
cscript //nologo GetRunningProcesses.vbs
cscript //nologo TestProcessRunning.vbs notepad.exe
cscript //nologo GetProcessMemoryUsage.vbs
```

The scripts list process names and PIDs, check for a process by exact name,
and report working-set and virtual-memory sizes. They do not terminate,
modify, or remotely control processes. They do not upload information.

Process termination, bulk termination, automatic restart or scheduling,
arbitrary command execution, elevation, remote control, software
installation/removal, stress testing, hidden-process inspection, and memory
dumps are intentionally not included.

Additional read-only diagnostics include per-process CPU counters (which can
exceed 100% across cores) and process creation times. Service exit codes and
pending states can also be inspected. `ExportServiceInventory.vbs
<new-csv-path>` writes only a local CSV report and refuses to overwrite an
existing file; it does not change service configuration.

Other individual queries report process owners, image paths, parent PIDs,
thread/handle counts, priorities, sessions, I/O counters, CPU times, memory
peaks/private memory, service-to-process mappings, descriptions, and service
dependency references. These details can expose local usernames, paths, or
session information; use them only on systems you administer.

## Windows service information scripts

These service queries are read-only and operate on the local computer:

```text
cscript //nologo GetWindowsServices.vbs
cscript //nologo GetServiceStatus.vbs Spooler
cscript //nologo GetServiceStartupType.vbs Spooler
cscript //nologo CheckMiddlewareServices.vbs
cscript //nologo WaitForServiceRunning.vbs Spooler 60
```

They list services, query one service by its exact service name, check common
SQL Server/IIS service states, and wait up to 300 seconds for a service to
reach `Running`. Invalid arguments, missing services, and timeouts are
reported with non-zero exit codes. They do not start, stop, reconfigure, or
remotely control services.

Service startup/recovery/account changes, remote operations, service
installation/removal, security-service enforcement, and automatic recovery
or maintenance actions are intentionally not included.

## Additional local process examples (No. 71-170)

The numbered scripts below are separate and operate only on the local computer. Process restarts, forced termination, memory dumps, arbitrary hidden execution, GUI clicking/SendKeys/screen capture, remote execution/control, UAC bypass or credential handling, antivirus exclusions, security enforcement, persistence, self-deletion, shutdown cancellation, external HTTP, and mass deployment are not implemented. Some requested numbers therefore have no corresponding VBS file; they are listed as excluded below.

Implemented safe local examples:

| No. | Script | Behavior |
|---:|---|---|
| 71 | `071_ProcessWorkingSet.vbs` | 指定プロセスのWorking Set  |
| 72 | `072_ProcessPagefileUsage.vbs` | 指定プロセスのページファイル使用量  |
| 73 | `073_ProcessCpuTimes.vbs` | 指定プロセスのKernel/User累計CPU時間  |
| 74 | `074_ProcessThreadsHandles.vbs` | 指定プロセスのスレッド数とハンドル数  |
| 75 | `075_ProcessIoBytes.vbs` | 指定プロセスのI/O読み書き量  |
| 76 | `076_ProcessSessionId.vbs` | 指定プロセスのセッションID  |
| 77 | `077_ProcessOwner.vbs` | 指定プロセスの所有者  |
| 78 | `078_ProcessExecutablePath.vbs` | 指定プロセスの実行ファイルパス  |
| 79 | `079_ProcessSignature.vbs` | プロセス実行ファイルのAuthenticode署名状態  |
| 80 | `080_ProcessFileVersion.vbs` | 指定プロセスの製品/ファイルバージョン  |
| 81 | `081_ProcessCreationTime.vbs` | 指定プロセスの起動日時  |
| 82 | `082_HighCpuMonitor.vbs` | CPU高負荷の有界サンプリング  |
| 83 | `083_MemoryTrend.vbs` | プロセスメモリの有界サンプリング  |
| 84 | `084_ProcessOpenHandleCount.vbs` | 指定プロセスのハンドル数（ハンドル名ではありません）  |
| 85 | `085_ProcessTree.vbs` | 親子プロセス関係のツリー表示  |
| 86 | `086_WatchProcessExit.vbs` | プロセス終了の期限付き監視（再起動なし）  |
| 87 | `087_DuplicateProcessCheck.vbs` | 重複起動の検出のみ  |
| 88 | `088_ActivateExistingWindow.vbs` | 既存ウィンドウのアクティブ化  |
| 89 | `089_TempCleanupPreview.vbs` | TEMP古いファイルの削除プレビューのみ  |
| 90 | `090_MaintenanceWindowCheck.vbs` | 保守時間帯の判定のみ  |
| 91 | `091_UpdatePreflight.vbs` | 更新前の関連プロセス確認  |
| 92 | `092_ProcessSequencePreview.vbs` | 起動順序のプレビューのみ  |
| 95 | `095_WorkstationLockEvents.vbs` | 画面ロックイベント照会  |
| 96 | `096_WorkstationUnlockEvents.vbs` | 画面ロック解除イベント照会  |
| 97 | `097_ProcessConcurrencyCount.vbs` | 同時プロセス数の取得  |
| 98 | `098_MemoryExhaustionEvents.vbs` | メモリ不足イベントの取得（ダンプなし）  |
| 99 | `099_NotRespondingProcesses.vbs` | 応答なしプロセス検出のみ  |
| 100 | `100_ProcessSummary.vbs` | プロセス集計サマリー  |
| 102 | `102_LaunchNotepadMaximized.vbs` | Notepad最大化起動の固定例  |
| 103 | `103_LaunchNotepadMinimized.vbs` | Notepad最小化起動の固定例  |
| 106 | `106_ActivateWindowTitle.vbs` | ウィンドウタイトルによるアクティブ化  |
| 122 | `122_LocalTcpConnections.vbs` | ローカルTCP接続の一覧  |
| 124 | `124_VpnAdapterStatus.vbs` | VPNアダプター状態照会のみ  |
| 126 | `126_RdpClientStatus.vbs` | ローカルRDPクライアント状態  |
| 130 | `130_PrintSpoolerStatus.vbs` | スプーラーサービス状態照会  |
| 135 | `135_TempProcessAudit.vbs` | TEMP配下の実行中プロセス監査  |
| 136 | `136_AuditProcessAllowlist.vbs` | プロセス許可リスト照合のみ  |
| 138 | `138_TempProcessStartAudit.vbs` | TEMP配下プロセス起動情報照会  |
| 143 | `143_SmartCardReaderStatus.vbs` | スマートカードリーダー状態照会  |
| 145 | `145_SandboxFeatureStatus.vbs` | Windows Sandbox機能状態照会のみ  |
| 146 | `146_RunFixedCommand.vbs` | 固定コマンドの同期/非同期呼び出し例  |
| 147 | `147_CaptureFixedCommandOutput.vbs` | 固定コマンドの出力/エラー取得例  |
| 148 | `148_PowerShellDateOutput.vbs` | 固定PowerShell照会例  |
| 151 | `151_ScheduledTaskStatus.vbs` | タスクスケジューラ状態照会のみ  |
| 152 | `152_ExitCodeBranchDemo.vbs` | 終了コード分岐デモ  |
| 153 | `153_ProcessTimingDemo.vbs` | 固定コマンドの実行時間計測  |
| 154 | `154_SingleInstanceLock.vbs` | ロックファイルの単一起動デモ  |
| 156 | `156_ProcessEnvironmentDemo.vbs` | プロセススコープ環境変数のデモ  |
| 157 | `157_CliArgumentsDemo.vbs` | コマンドライン引数表示デモ  |
| 158 | `158_DroppedFileMetadata.vbs` | ドロップファイルのメタデータ取得  |
| 160 | `160_ConsoleProgressDemo.vbs` | コンソール進捗表示デモ  |
| 161 | `161_LockFileStatus.vbs` | ロックファイル状態照会  |
| 162 | `162_VerboseModeDemo.vbs` | verboseモード切替デモ  |
| 164 | `164_ShutdownEventHistory.vbs` | シャットダウン履歴照会のみ  |
| 167 | `167_AppendProcessLog.vbs` | プロセス数のローカルログ追記  |
| 169 | `169_WaitForCpuBelowThreshold.vbs` | CPU負荷が下がるまでの期限付き待機  |
| 170 | `170_ProcessDashboard.vbs` | プロセス一覧のローカルHTML出力  |

Intentionally excluded numbers: 93, 94, 101, 104, 105, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 123, 125, 127, 128, 129, 131, 132, 133, 134, 137, 139, 140, 141, 142, 144, 149, 150, 155, 159, 163, 165, 166, 168. These items require process termination/restart, remote or privileged control, security-policy changes, data capture, persistence, or other operations outside this safe local scope.
