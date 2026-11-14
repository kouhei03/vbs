> **ファイル概要:** Windowsサービスの照会・監視スクリプトの一覧と制限事項。
> **実行方法:** この文書を参照し、Windowsのコマンドプロンプトでこのフォルダーから `cscript //nologo GetAllServices.vbs` を実行。
> **実行結果:** 例ではサービス名、表示名、状態がコンソールに表示されます。

# Windows service inspection and monitoring

Run each script locally with Windows Script Host. These scripts query service
configuration and status; none starts, stops, restarts, reconfigures, installs,
or removes a service.

Examples:

```text
cscript //nologo GetAllServices.vbs
cscript //nologo GetServiceStatus.vbs Spooler
cscript //nologo WaitForServiceRunning.vbs Spooler 60
cscript //nologo WatchServiceHealth.vbs Spooler 5 60
cscript //nologo ExportServiceInventory.vbs C:\Temp\services.csv
cscript //nologo GetServiceCrashCounts.vbs 7
cscript //nologo WatchSecurityServiceStatus.vbs
```

`WatchServiceHealth.vbs` is a bounded, console-only status monitor; it does
not attempt recovery. CSV export creates a new file and refuses to overwrite
an existing one. Security-service status reports only local state.

Other read-only scripts check a service's existence, startup mode, dependency
references, logon account name, exit codes, and common SQL Server/IIS states.
The crash count report counts System log events 7031 and 7034 over a 1-30 day
window without displaying event message text.

Service lifecycle changes, automatic recovery, remote control, service
installation/removal, account changes, event-log writes, and system policy
changes are intentionally excluded.

## Additional local service diagnostics (No. 106-205)

These numbered scripts perform local, read-only inspection. The CSV/HTML report scripts create a new local file and refuse to overwrite existing files. They do not start, stop, restart, reconfigure, install, or remove services; they do not remotely control hosts or transmit data. Recovery action and service ACL queries may reveal sensitive local configuration, so use them only on systems you administer.

Implemented scripts:

| No. | Script | Behavior   |
|---:|---|---  |
| 106 | `106_DependentServices.vbs` | 依存しているサービスの列挙   |
| 107 | `107_DependencyStartOrderPreview.vbs` | 必要サービスの起動順プレビュー（実行なし）   |
| 108 | `108_CascadeStopImpactPreview.vbs` | 停止時に影響を受けるサービスの状態プレビュー（停止なし）   |
| 109 | `109_DelayedAutoStartStatus.vbs` | 遅延自動開始レジストリ値の読み取り   |
| 111 | `111_ServiceRecoveryActions.vbs` | サービス回復アクション設定の読み取り   |
| 112 | `112_ServiceFailureProgram.vbs` | 失敗時に登録された回復プログラム設定の読み取り   |
| 113 | `113_ServiceFailureResetPeriod.vbs` | 失敗回数リセット期間の読み取り   |
| 114 | `114_ServiceFailureRestartDelay.vbs` | 回復再起動遅延設定の読み取り   |
| 115 | `115_ServiceDesktopInteractStatus.vbs` | DesktopInteract設定の読み取り   |
| 117 | `117_ServiceLogonAccountAudit.vbs` | ログオンアカウント名の読み取り   |
| 118 | `118_ServiceSidTypeStatus.vbs` | サービスSID種別の読み取り   |
| 120 | `120_ServiceConfigurationReadOnly.vbs` | 基本サービス構成情報の読み取り   |
| 121 | `121_ServicesPipeTimeoutStatus.vbs` | ServicesPipeTimeout値の読み取り   |
| 122 | `122_DisabledServicesRestorePreview.vbs` | 無効サービスの復旧検討用一覧   |
| 123 | `123_ServiceGroupInventory.vbs` | サービスグループ情報の一覧   |
| 125 | `125_ServiceDisplayNames.vbs` | サービス名・表示名の一覧   |
| 146 | `146_SqlServerServiceStatus.vbs` | SQL Server関連サービスの状態確認   |
| 147 | `147_IisServiceStatus.vbs` | IIS関連サービスの状態確認   |
| 148 | `148_DirectoryServicesStatus.vbs` | AD/DNS関連サービスの状態確認   |
| 149 | `149_DhcpDnsServiceStatus.vbs` | DHCP/DNSサービスの状態確認   |
| 150 | `150_PrintSpoolerStatus.vbs` | Print Spoolerの状態確認のみ   |
| 151 | `151_WindowsUpdateServicesStatus.vbs` | Windows Update関連サービスの状態確認のみ   |
| 152 | `152_SecurityServicesStatus.vbs` | ローカルセキュリティサービスの状態確認   |
| 153 | `153_BackupVssStatus.vbs` | バックアップ/VSS関連サービスの状態確認   |
| 154 | `154_MailServiceStatus.vbs` | メール関連サービスの状態確認   |
| 155 | `155_DfsServiceStatus.vbs` | DFS関連サービスの状態確認   |
| 156 | `156_RemoteDesktopStatus.vbs` | RDP関連サービスの状態確認のみ   |
| 157 | `157_TaskSchedulerServiceStatus.vbs` | タスクスケジューラサービスの状態確認のみ   |
| 158 | `158_EventLogServiceStatus.vbs` | イベントログサービスの状態確認のみ   |
| 159 | `159_WindowsTimeStatus.vbs` | Windows Timeサービスの状態確認のみ   |
| 160 | `160_CertificateServiceStatus.vbs` | 証明書サービスの状態確認   |
| 161 | `161_QueryServiceDetails.vbs` | 指定サービスの詳細読み取り   |
| 163 | `163_MessageQueueServiceStatus.vbs` | MSMQサービスの状態確認   |
| 164 | `164_VirtualizationServiceStatus.vbs` | Hyper-V関連サービスの状態確認   |
| 165 | `165_MonitoringAgentStatus.vbs` | 監視エージェントの状態確認   |
| 168 | `168_PrivilegedServiceAccounts.vbs` | 組み込みサービスアカウント使用サービスの確認   |
| 169 | `169_ServiceAllowlistAudit.vbs` | 許可リストとサービス名の読み取り専用比較   |
| 170 | `170_ServiceBinarySignatureStatus.vbs` | サービス実行ファイルのローカル署名確認   |
| 171 | `171_SecurityProductStatus.vbs` | セキュリティ製品関連サービスの状態確認   |
| 172 | `172_ServiceAclQuery.vbs` | サービスセキュリティ記述子の読み取り   |
| 173 | `173_ServiceAuditPolicyStatus.vbs` | 監査ポリシー状態の読み取り   |
| 181 | `181_ServiceStateSnapshot.vbs` | 新規CSVへのサービス状態スナップショット   |
| 185 | `185_ServiceAuditHtml.vbs` | 新規HTML監査レポート   |
| 186 | `186_BoundedServicePoll.vbs` | 期限付きサービス状態ポーリング   |
| 187 | `187_ServiceBackoffObservation.vbs` | 段階的間隔での状態観察のみ   |
| 193 | `193_ServiceExitCodeGuide.vbs` | 終了コードと一般的な説明の表示   |
| 195 | `195_DebugModeArgumentDemo.vbs` | 詳細表示引数の安全な解析例   |
| 196 | `196_ServiceQueryErrorHandling.vbs` | WMI照会エラー処理の例   |
| 199 | `199_ThreeFailureStatusFilter.vbs` | 連続3回の状態確認フィルターのみ   |
| 201 | `201_ServiceDashboard.vbs` | 新規HTMLサービスダッシュボード   |
| 202 | `202_LowOverheadPollDemo.vbs` | 低頻度・回数限定のサービス照会   |
| 204 | `204_WmiServiceCompatibility.vbs` | WMIクラス利用可否の互換性確認   |
| 205 | `205_ServiceInventorySummary.vbs` | 状態・スタートアップ種別の集計   |

Intentionally excluded numbers: 110, 116, 119, 124, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144, 145, 162, 166, 167, 174, 175, 176, 177, 178, 179, 180, 182, 183, 184, 188, 189, 190, 191, 192, 194, 197, 198, 200, 203. These requests involve service/account/configuration changes, remote management, automatic recovery or escalation, security-policy enforcement, external notifications/distribution, or other actions outside this local read-only scope.
