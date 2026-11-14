# ネットワーク・共有・Web・リモート管理 (No. 1-50)

Windows Script Host用のVBScript例です。Windowsのコマンドプロンプトから
`cscript //nologo <script.vbs> [arguments]` の形式で個別に実行してください。
各スクリプトの先頭にUsageを記載しています。

| No. | スクリプト | 内容 |
|---:|---|---|
| 1 | `001_NetworkDriveAudit.vbs` | 現在のネットワークドライブ一覧（接続変更なし） |
| 2 | `002_NetworkDriveDisconnectPreview.vbs` | 指定ドライブの接続状態確認（切断なし） |
| 3 | `003_PingSingleHost.vbs` | 単一ホストへの上限付きPing |
| 4 | `004_LocalIpConfiguration.vbs` | ローカルIP・サブネット・ゲートウェイ表示 |
| 5 | `005_DnsCacheSummary.vbs` | DNSキャッシュの件数概要（詳細表示・消去なし） |
| 6 | `006_HttpStatusProbe.vbs` | 単一URLへのHTTP GETとステータス確認 |
| 7 | `007_LocalHtmlSearch.vbs` | ローカルHTMLファイルのキーワード検索 |
| 8 | `008_DefaultPrinterAudit.vbs` | 既定プリンター確認 |
| 9 | `009_FirewallProfileAudit.vbs` | Defender Firewallプロファイル確認 |
| 10 | `010_DnsLookup.vbs` | 単一ホストのDNS名前解決 |
| 11 | `011_NetworkAdapterStatus.vbs` | ローカルNIC状態照会 |
| 12 | `012_ProxySettingsAudit.vbs` | 現在のユーザープロキシ設定確認 |
| 13 | `013_PublicIpLookupExcluded.vbs` | 外部IPチェッカーへの接続を行わない案内 |
| 14 | `014_ShareAclAudit.vbs` | 単一共有パスのACL照会 |
| 15 | `015_LocalOsInfo.vbs` | ローカルOS情報表示 |
| 16 | `016_LocalServiceAudit.vbs` | 単一ローカルサービス状態照会 |
| 17 | `017_VpnAdapterAudit.vbs` | VPN関連アダプター状態確認 |
| 18 | `018_PingLatencyReport.vbs` | 単一ホストへの限定Ping結果を新規ログへ保存 |
| 19 | `019_LocalJsonValidation.vbs` | ローカルJSONファイルの構文検証 |
| 20 | `020_LocalAlertLog.vbs` | SMTP送信の代替となるローカルアラート記録 |
| 21 | `021_SharedFileMetadataAudit.vbs` | 単一共有ファイルのメタデータとACL確認 |
| 22 | `022_WifiAdapterAudit.vbs` | Wi-Fiアダプター状態照会（プロファイル変更なし） |
| 23 | `023_CheckSingleHost.vbs` | 単一ホストへの1回Ping |
| 24 | `024_LocalTransferManifest.vbs` | 転送前のローカルファイルSHA-256表示 |
| 25 | `025_TimeServiceAudit.vbs` | Windows Timeサービス状態照会 |
| 26 | `026_LocalSnmpServiceAudit.vbs` | ローカルSNMPサービス状態照会 |
| 27 | `027_PrintQueueAudit.vbs` | 印刷キュー照会（ジョブ削除なし） |
| 28 | `028_BrowserCacheSizeAudit.vbs` | 代表的キャッシュのサイズ確認（削除なし） |
| 29 | `029_RdpConfigurationAudit.vbs` | RDP関連レジストリとFirewall状態照会 |
| 30 | `030_VpnTunnelStatusAudit.vbs` | VPN関連アダプター状態照会 |
| 31 | `031_IpConfigurationSnapshot.vbs` | IP設定を新規ローカルスナップショットに保存 |
| 32 | `032_ProxyCredentialCacheNotice.vbs` | 資格情報を扱わない安全な案内 |
| 33 | `033_WebResponseTime.vbs` | 単一URLの応答時間・HTTP状態確認 |
| 34 | `034_NicPowerStatusAudit.vbs` | NIC電源管理プロパティ確認 |
| 35 | `035_IpProtocolBindingAudit.vbs` | IPv4/IPv6アドレス構成の確認 |
| 36 | `036_ShareAccessChecklist.vbs` | 単一UNC共有へのアクセス確認と診断項目表示 |
| 37 | `037_LocalNetworkReport.vbs` | ローカルIP・MAC情報の新規レポート |
| 38 | `038_NetworkFileCopyAudit.vbs` | 2ファイルのサイズ・SHA-256比較（コピーなし） |
| 39 | `039_DomainMembershipAudit.vbs` | ローカルのドメイン参加状態確認 |
| 40 | `040_PrinterDriverPortAudit.vbs` | プリンタードライバー・TCPポート一覧 |
| 41 | `041_RemoteShutdownPreview.vbs` | リモート停止コマンドを実行しないプレビュー |
| 42 | `042_LocalTransferChecksum.vbs` | ローカルファイルのチェックサム一覧作成 |
| 43 | `043_ApprovedUrlList.vbs` | URLリストの検証・表示（ブラウザー起動なし） |
| 44 | `044_RouteTableAudit.vbs` | ローカルルーティングテーブル表示 |
| 45 | `045_MacInventory.vbs` | 有効NICのMACアドレスを新規CSVへ出力 |
| 46 | `046_NicMetricAudit.vbs` | NICメトリック照会（変更なし） |
| 47 | `047_ExpiringCertificatesAudit.vbs` | ローカル証明書ストアの90日以内期限切れ監査 |
| 48 | `048_LocalAddressCount.vbs` | ローカルNICのIPv4アドレス数集計（ホスト走査なし） |
| 49 | `049_ChecksumCompare.vbs` | 2ファイルのSHA-256比較 |
| 50 | `050_LocalNetworkDiagnostics.vbs` | ローカル情報を新規ZIPに収集 |

## 制限・注意

- リモートマシン制御、セグメントスキャン、資格情報の取得・削除、Wi-Fiプロファイル操作、NIC切替、Firewall/RDP設定変更、権限変更、プリンターキュー／ドライバー削除、リモートシャットダウン、FTP/クラウド転送は実装していません。該当項目はローカル監査、単一対象の限定確認、またはプレビューに置き換えています。
- Web通信を行う例は指定された単一URLへのGETだけです。認証情報やPOSTデータは送りません。URLは信頼できる対象に限ってください。
- PowerShellを利用する例ではコマンドを固定し、任意のコマンド文字列を受け取りません。WMI、PowerShell、`certutil`、Windows Firewall CIMの有無・権限によっては照会に失敗します。
- ネットワーク診断には内部IP、MAC、ルート情報が含まれる場合があります。生成されたレポートやZIPを共有する前に内容を確認してください。
- No. 19はWindows PowerShellの`ConvertFrom-Json`を使います。No. 50はWindows Shell ZIP機能を使います。
