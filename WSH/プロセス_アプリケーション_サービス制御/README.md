# プロセス・アプリケーション・サービス制御 (No. 51-100)

Windows Script Host用のVBScript例です。Windowsのコマンドプロンプトから
`cscript //nologo <script.vbs> [arguments]` の形式で個別に実行してください。
各スクリプトの冒頭にUsageを記載しています。

| No. | スクリプト | 内容 |
|---:|---|---|
| 51 | `051_CheckProcessRunning.vbs` | 指定プロセスの死活確認 |
| 52 | `052_HungProcessAudit.vbs` | 応答なしプロセスの監査（終了なし） |
| 53 | `053_OpenTextFileInNotepad.vbs` | 指定ファイルを固定のNotepadで開く |
| 54 | `054_ServiceStatusAudit.vbs` | サービス状態照会（制御なし） |
| 55 | `055_ServiceStartupTypeAudit.vbs` | サービス起動種類の照会 |
| 56 | `056_RunFixedCommandSync.vbs` | 固定コマンドの同期実行 |
| 57 | `057_SequentialNotepadLaunch.vbs` | Notepadの有界な順次起動 |
| 58 | `058_ProcessPresenceAudit.vbs` | 指定プロセス候補の監査（停止なし） |
| 59 | `059_ProcessResourceSnapshot.vbs` | CPUカウンターとprivate working setのスナップショット |
| 60 | `060_CrashEventAudit.vbs` | Application Errorイベント監査（再起動なし） |
| 61 | `061_ProcessSessionAudit.vbs` | プロセスのSessionId照会 |
| 62 | `062_VisibleFixedCommand.vbs` | 固定コマンドの可視実行例 |
| 63 | `063_ExitCodeBranch.vbs` | 固定コマンドの終了コード分岐 |
| 64 | `064_InstallerFileAudit.vbs` | インストーラーファイル情報の確認（実行なし） |
| 65 | `065_ActivateWindowTitle.vbs` | 既存ウィンドウのアクティブ化 |
| 66 | `066_InteractiveLoginAudit.vbs` | 対話ログオンユーザー確認 |
| 67 | `067_ProcessTreeAudit.vbs` | 指定プロセスの親PID監査（終了なし） |
| 68 | `068_UpdatePreflightAudit.vbs` | バージョン・更新日時の事前確認（更新なし） |
| 69 | `069_SingleInstanceLockDemo.vbs` | 一時ロックファイルによる単一起動デモ |
| 70 | `070_AdminTokenAudit.vbs` | 管理者ロール照会（昇格なし） |
| 71 | `071_ServiceDependencyAudit.vbs` | サービス依存関係の照会（制御なし） |
| 72 | `072_ProcessPriorityAudit.vbs` | プロセス優先度の照会（変更なし） |
| 73 | `073_MaintenanceProcessPreview.vbs` | メンテナンス候補の表示（終了なし） |
| 74 | `074_CrashEventMetadata.vbs` | クラッシュイベントのメタデータ照会（ダンプなし） |
| 75 | `075_LocalTextPipeline.vbs` | 外部コマンドを使わないテキスト処理パイプライン |
| 76 | `076_MiddlewareServiceAudit.vbs` | 代表的ミドルウェアサービスの状態照会 |
| 77 | `077_StartupCommandAudit.vbs` | WMIスタートアップ項目の監査 |
| 78 | `078_FixedComObjectDemo.vbs` | FileSystemObjectの固定COM呼び出し例 |
| 79 | `079_AntivirusStatusAudit.vbs` | SecurityCenter2の製品状態照会 |
| 80 | `080_PrintSpoolerAudit.vbs` | スプーラー状態照会（再起動なし） |
| 81 | `081_HyperVGuestStateAudit.vbs` | Hyper-V仮想マシン状態照会（制御なし） |
| 82 | `082_AuthenticodeAudit.vbs` | Authenticode署名状態の照会 |
| 83 | `083_ProcessUseBySession.vbs` | SessionId別プロセス数の集計 |
| 84 | `084_CaptureFixedStdout.vbs` | 固定コマンドの標準出力を新規ログへ保存 |
| 85 | `085_FixedStderrAudit.vbs` | 固定デモの標準エラー検知（通知送信なし） |
| 86 | `086_FixedCommandTimeoutAudit.vbs` | 固定コマンドの期限監視（強制終了なし） |
| 87 | `087_ShutdownPreviewOnly.vbs` | シャットダウン条件の確認のみ |
| 88 | `088_BoundedSleepDemo.vbs` | 上限付きSleepポーリング例 |
| 89 | `089_ProcessEnvironmentDemo.vbs` | プロセス環境変数と作業ディレクトリの例 |
| 90 | `090_VerboseModeDemo.vbs` | normal/debugログ切替例 |
| 91 | `091_WindowTitleAudit.vbs` | ウィンドウタイトルの確認（SendKeysなし） |
| 92 | `092_ExitCodeMessageMap.vbs` | 終了コードから定義済みメッセージへの対応づけ |
| 93 | `093_WaitForServiceRunning.vbs` | サービスがRunningになるまで最大60秒待機 |
| 94 | `094_ProcessMemoryThresholdAudit.vbs` | メモリ閾値監査（再起動なし） |
| 95 | `095_LogonSessionAudit.vbs` | ログオンセッション種別の列挙 |
| 96 | `096_SessionLockStatusAudit.vbs` | 対話セッション情報の表示（ロックなし） |
| 97 | `097_ConsoleProgressDemo.vbs` | コンソール進捗表示デモ |
| 98 | `098_FixedBatVbsInteropDemo.vbs` | 固定の`ver`照会によるBAT/VBS連携デモ |
| 99 | `099_StartupRegistryAudit.vbs` | Runキー登録値の監査（変更なし） |
| 100 | `100_ReadOnlyMaintenanceLauncher.vbs` | プロセス数・代表サービス状態の読み取り専用サマリー |

## 制限・注意

- これらはローカルの診断・デモ用です。任意コマンド実行、リモート操作、強制終了、プロセス／サービスの自動再起動、サービス構成変更、サイレントインストール、セキュリティ製品の停止、ダンプ取得、シャットダウン、SendKeys、レジストリ永続化変更は実装していません。該当する例は読み取り専用の監査またはプレビューに置き換えています。
- PowerShellを利用する例も、コマンドはスクリプト内で固定しています。ユーザー入力をコマンド文字列として実行しません。
- `065_ActivateWindowTitle.vbs`と`091_WindowTitleAudit.vbs`は一致するウィンドウをアクティブ化する場合があります。`069_SingleInstanceLockDemo.vbs`は異常終了時にTEMP内のロックファイルが残ることがあります。実行中の別インスタンスがないことを確認してから手動で処理してください。
- Windows上での実行を想定しています。WMIクラス、Hyper-V、SecurityCenter2、PowerShellの有無やアクセス権により、照会できない項目があります。
