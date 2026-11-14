# レジストリ・環境変数・OS設定・セキュリティ自動化 (No. 151-200)

Windows Script Host用のVBScript例です。Windowsのコマンドプロンプトから
`cscript //nologo <script.vbs> [arguments]` の形式で個別に実行してください。
各スクリプトの先頭にUsageを記載しています。

| No. | スクリプト | 内容 |
|---:|---|---|
| 151 | `151_RegistryValueAudit.vbs` | HKCU\Software配下の値を読み取り |
| 152 | `152_DemoRegistryKeyAudit.vbs` | デモ用レジストリキーの存在確認 |
| 153 | `153_RegistryCleanupPreview.vbs` | デモキーの削除候補表示（削除なし） |
| 154 | `154_EnvironmentAudit.vbs` | USER/SYSTEM/PROCESS環境変数の監査（秘密名をマスク） |
| 155 | `155_PathAppendPreview.vbs` | PATH追加の重複確認・プレビュー |
| 156 | `156_EnvironmentDeletePreview.vbs` | 環境変数の削除候補確認（削除なし） |
| 157 | `157_AutoLogonAudit.vbs` | 自動ログオン状態確認（パスワード値は非読取） |
| 158 | `158_UsbStoragePolicyAudit.vbs` | USBSTORポリシー値の確認 |
| 159 | `159_WindowsUpdatePolicyAudit.vbs` | Windows Updateポリシー確認 |
| 160 | `160_TaskManagerPolicyAudit.vbs` | Task Manager制限ポリシー確認 |
| 161 | `161_ControlPanelPolicyAudit.vbs` | Control Panel制限ポリシー確認 |
| 162 | `162_WallpaperPolicyAudit.vbs` | 壁紙設定値確認 |
| 163 | `163_ScreensaverPolicyAudit.vbs` | スクリーンセーバー設定確認 |
| 164 | `164_FileAssociationAudit.vbs` | 拡張子のUserChoice ProgId照会 |
| 165 | `165_ContextMenuAudit.vbs` | ユーザーコンテキストメニュー値の照会 |
| 166 | `166_FirstRunFlagAudit.vbs` | デモ用初回起動フラグの照会 |
| 167 | `167_UacPolicyAudit.vbs` | UAC EnableLUA値の照会 |
| 168 | `168_ExplorerDisplayAudit.vbs` | 隠しファイル・拡張子表示設定確認 |
| 169 | `169_IfeoAudit.vbs` | notepad.exeのIFEO Debugger値確認 |
| 170 | `170_GroupPolicyResultAudit.vbs` | 適用済みユーザーポリシー照会 |
| 171 | `171_RegistryNotificationNotice.vbs` | 設定変更・通知を行わない案内 |
| 172 | `172_UserHiveInventory.vbs` | プロファイルとNTUSER.DAT存在確認（編集なし） |
| 173 | `173_FastStartupAudit.vbs` | 高速スタートアップ関連値確認 |
| 174 | `174_NotificationPolicyAudit.vbs` | 通知設定値確認 |
| 175 | `175_ClassicContextMenuAudit.vbs` | クラシックメニュー上書き値確認 |
| 176 | `176_BrowserPolicyAudit.vbs` | Edge/Chromeポリシー設定の有無確認 |
| 177 | `177_CustomApplicationEvent.vbs` | Applicationログへ固定の情報イベントを1件記録 |
| 178 | `178_PasswordPolicyAudit.vbs` | ローカルアカウントのパスワード変更ポリシー確認 |
| 179 | `179_AccountLockoutPolicyAudit.vbs` | アカウントロックアウトポリシー照会 |
| 180 | `180_GuestAccessPolicyAudit.vbs` | ゲスト／匿名アクセス関連値確認 |
| 181 | `181_NicPowerManagementAudit.vbs` | NIC電源管理プロパティ確認 |
| 182 | `182_LicenseStorageNotice.vbs` | ライセンス情報を保存しない安全案内 |
| 183 | `183_UninstallKeyAudit.vbs` | アンインストール登録の件数のみ確認 |
| 184 | `184_RegistryTemplatePreview.vbs` | HKCUデモ用.regテンプレートを新規作成（インポートなし） |
| 185 | `185_DemoRegistryVerification.vbs` | デモ用レジストリ値の事後確認 |
| 186 | `186_SafeEnvironmentSnapshotCompare.vbs` | PATH/TEMP/USERPROFILEのスナップショット比較 |
| 187 | `187_UserProfileAudit.vbs` | プロファイル件数とロード済み件数 |
| 188 | `188_PrinterDriverRegistryAudit.vbs` | プリンタードライバー登録の読み取り監査 |
| 189 | `189_NetworkPolicyAudit.vbs` | TCP/IP関連値の監査（ロールバックなし） |
| 190 | `190_TimeZoneAudit.vbs` | ローカルタイムゾーン情報確認 |
| 191 | `191_UsbStorageHistoryCount.vbs` | USBSTOR履歴キー件数（識別子非表示） |
| 192 | `192_UsbHistoryCleanupNotice.vbs` | USB履歴を削除しない案内 |
| 193 | `193_RegistryErrorHandlerDemo.vbs` | レジストリ読取エラー処理デモ |
| 194 | `194_ConsoleProgressDemo.vbs` | コンソール進捗表示デモ |
| 195 | `195_WindowsCompatibilityAudit.vbs` | Windowsバージョン・アーキテクチャ確認 |
| 196 | `196_ReadOnlyAdminMenu.vbs` | 対話式の読み取り専用監査メニュー |
| 197 | `197_VisualEffectsPolicyAudit.vbs` | 視覚効果プリセット値確認 |
| 198 | `198_DefenderPolicyAudit.vbs` | Defenderクラウド保護等の状態照会 |
| 199 | `199_FirewallNotificationAudit.vbs` | Firewall通知状態確認 |
| 200 | `200_ReadOnlyEnvironmentLauncher.vbs` | OS・環境・UACの読み取り専用サマリー |

## 制限・注意

- レジストリ値・環境変数・OS設定の書込み／削除、ポリシー変更、gpupdate、ハイブのロード・編集、設定の一括配布、UAC・Defender・Firewallの弱体化は実装していません。該当項目は読み取り専用の監査、プレビュー、または案内に置き換えています。
- No. 177だけは要求内容に合わせてApplicationログへ固定の情報イベントを1件記録します。Security監査ログには書き込みません。
- No. 184はHKCUの専用デモキーを対象にした.regファイルを作るだけです。自動インポートや一斉適用はしません。
- No. 154は環境変数を表示しますが、PASSWORD/SECRET/TOKEN/CREDENTIAL/API_KEYを含む変数名の値はマスクします。No. 157はAutoLogonのパスワード値を読み取りません。
- 管理者権限、WMI、PowerShell、ローカルポリシーによって照会結果が得られない場合があります。Windows上で、対象端末の管理権限とプライバシー要件を確認して使用してください。
