# Registry and environment inspection scripts
<!-- 解説: この見出しは「Registry and environment inspection scripts」を説明する節のタイトルです。 -->

The scripts in this directory are read-only. They can read a string or DWORD
registry value, check a key, enumerate a key's immediate subkeys, or display
only the standard process paths `SystemRoot`, `TEMP`, and `TMP`.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: The scripts in this directory are read-only. They can read a string or DWORD -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: registry value, check a key, enumerate a key's immediate subkeys, or display -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: only the standard process paths `SystemRoot`, `TEMP`, and `TMP`. -->

Examples:
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Examples: -->

```text
cscript //nologo ReadRegistryString.vbs HKCU "Control Panel\Desktop" Wallpaper
cscript //nologo ReadRegistryDword.vbs HKCU "Software\Example" Setting
cscript //nologo TestRegistryKey.vbs HKCU "Software\Example"
cscript //nologo ListRegistrySubkeys.vbs HKCU Software
cscript //nologo GetStandardEnvironmentPaths.vbs
```
<!-- 解説: この行はREADME内の実行例コードブロックの開始を示します。 -->
<!-- 解説: このコマンド例はReadRegistryString.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はReadRegistryDword.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はTestRegistryKey.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はListRegistrySubkeys.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はGetStandardEnvironmentPaths.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: この行は実行例コードブロックの終端です。 -->

Registry hives accepted by the registry scripts are `HKCU`, `HKLM`, `HKCR`,
`HKU`, and `HKCC`. Registry writes, deletion, import, repair, and persistent
environment-variable changes are intentionally not included. Autologon,
policy/security, proxy/RDP/UAC, IFEO, and profile-hive modifications are also
excluded; read-only state checks for selected items are available below.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Registry hives accepted by the registry scripts are `HKCU`, `HKLM`, `HKCR`, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: `HKU`, and `HKCC`. Registry writes, deletion, import, repair, and persistent -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: environment-variable changes are intentionally not included. Autologon, -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: policy/security, proxy/RDP/UAC, IFEO, and profile-hive modifications are also -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: excluded; read-only state checks for selected items are available below. -->

## Additional read-only diagnostics (items 107-140)
<!-- 解説: この見出しは、追加した読み取り専用診断スクリプトを一覧化する節を示します。 -->

The following scripts provide inspection or preview equivalents for the requested
items. They do not write, delete, import, or repair registry data. Export items
only check the source; environment and policy items report current state. Items
that require a state change intentionally remain out of scope.
<!-- 解説: この段落は追加スクリプトが読み取り専用であり、状態変更操作は含めないことを説明しています。 -->

| No. | Script | Read-only behavior |
|---:|---|---|
| 107 | `PreviewRegistryStringWrite.vbs` | Validates and previews the target and proposed string length without writing. | <!-- 解説: No.107はREG_SZ書き込み要求の対象を表示するだけで、レジストリを変更しません。 -->
| 109 | `InspectMultiStringAndBinary.vbs` | Reads REG_MULTI_SZ values or reports binary size without exposing binary data. | <!-- 解説: No.109は複数行文字列またはバイナリ値を読み取り、バイナリ本体は表示しません。 -->
| 111 | `InspectRegistryDeleteTarget.vbs` | Lists value names and types under a target key; nothing is deleted. | <!-- 解説: No.111は削除対象キーの値名と型を確認し、削除は行いません。 -->
| 113 | `CheckRegistryExportSource.vbs` | Checks whether the requested key is readable; does not create an export file. | <!-- 解説: No.113はエクスポート元キーの読み取り可否だけを確認します。 -->
| 114 | `InspectRegFile.vbs` | Checks a REG file header and counts key sections; does not import it. | <!-- 解説: No.114はREGファイルの基本形式を確認し、インポートしません。 -->
| 115 | `ListSystemEnvironmentVariables.vbs` | Lists system environment variables without changing them. | <!-- 解説: No.115はシステム環境変数を一覧表示し、変更しません。 -->
| 116 | `ListUserEnvironmentVariables.vbs` | Lists current-user environment variables without changing them. | <!-- 解説: No.116はユーザー環境変数を一覧表示し、変更しません。 -->
| 117 | `InspectPathDuplicates.vbs` | Reports duplicate entries in the current process PATH. | <!-- 解説: No.117は現在のPATH内の重複候補を報告し、追加はしません。 -->
| 118 | `InspectEnvironmentRemovalTarget.vbs` | Checks whether a user environment variable exists; does not remove it. | <!-- 解説: No.118は環境変数削除対象の有無を確認し、削除や通知はしません。 -->
| 119 | `InspectAutoLogonState.vbs` | Reports AutoAdminLogon and user name; never reads a password. | <!-- 解説: No.119は自動ログオン状態を確認し、パスワード値にはアクセスしません。 -->
| 120 | `InspectUSBStoragePolicy.vbs` | Reads the USBSTOR startup value without changing the policy. | <!-- 解説: No.120はUSBストレージ関連の設定値を読み取ります。 -->
| 121 | `InspectWindowsUpdatePolicy.vbs` | Lists Windows Update policy value names and types. | <!-- 解説: No.121はWindows Updateポリシーの値名と型を確認します。 -->
| 122 | `InspectTaskManagerPolicy.vbs` | Reads the current user's Task Manager restriction value. | <!-- 解説: No.122はタスクマネージャー制限の現在値を確認します。 -->
| 123 | `InspectControlPanelPolicy.vbs` | Reads the current user's Control Panel restriction value. | <!-- 解説: No.123はコントロールパネル制限の現在値を確認します。 -->
| 124 | `InspectDesktopWallpaper.vbs` | Reports the current user's wallpaper path. | <!-- 解説: No.124は現在の壁紙設定を読み取ります。 -->
| 125 | `InspectScreenSaverSettings.vbs` | Reads screen-saver activation, timeout, and secure-lock settings. | <!-- 解説: No.125はスクリーンセーバーとロック関連の現在値を確認します。 -->
| 126 | `InspectBrowserHomePages.vbs` | Reads the IE user homepage and Edge managed homepage values. | <!-- 解説: No.126はブラウザーのホームページ関連設定を読み取ります。 -->
| 127 | `InspectFileAssociation.vbs` | Reads the extension's default and user-selected ProgID. | <!-- 解説: No.127は拡張子に関連付けられたProgIDを読み取ります。 -->
| 128 | `InspectContextMenuEntries.vbs` | Lists common shell context-menu registration subkeys. | <!-- 解説: No.128は代表的なコンテキストメニュー登録キーを列挙します。 -->
| 129 | `InspectApplicationFirstRunFlag.vbs` | Checks for a specified value without clearing or changing it. | <!-- 解説: No.129は指定された初回起動フラグの有無を確認します。 -->
| 130 | `InspectProxySettings.vbs` | Reads proxy settings and omits values that may contain credentials. | <!-- 解説: No.130はプロキシ設定を読み、資格情報を含む可能性がある値は伏せます。 -->
| 131 | `InspectRemoteDesktopPolicy.vbs` | Reads the RDP connection policy value. | <!-- 解説: No.131はリモートデスクトップ接続許可に関する現在値を確認します。 -->
| 132 | `InspectUACSettings.vbs` | Reads selected UAC policy values without changing them. | <!-- 解説: No.132はUAC関連設定を読み取り、変更しません。 -->
| 133 | `InspectExplorerDisplaySettings.vbs` | Reads hidden-file and file-extension display values. | <!-- 解説: No.133は隠しファイルと拡張子表示の設定値を確認します。 -->
| 134 | `InspectIFEOBlockSettings.vbs` | Checks whether a Debugger value exists; does not reveal its command data. | <!-- 解説: No.134はIFEOのDebugger値の有無だけを確認し、値データは表示しません。 -->
| 135 | `AuditLicenseState.vbs` | Reports licensing status without retrieving product keys. | <!-- 解説: No.135はプロダクトキーを取得せずライセンス状態を報告します。 -->
| 136 | `ShowAppliedGroupPolicies.vbs` | Runs the read-only `gpresult` report; does not force policy updates. | <!-- 解説: No.136は適用済みポリシーを照会し、強制更新は実行しません。 -->
| 137 | `CompareEnvironmentRefreshState.vbs` | Compares PATH lengths in registry and current process; sends no broadcast. | <!-- 解説: No.137は環境変数の現在値を比較し、OS通知APIは呼び出しません。 -->
| 138 | `InspectRegistryHiveFiles.vbs` | Reports standard hive file metadata; does not load or repair hives. | <!-- 解説: No.138はレジストリハイブファイルのメタデータを確認し、修復しません。 -->
| 139 | `InspectUserProfileHives.vbs` | Lists profile SID, loaded state, and NTUSER.DAT presence; does not mount hives. | <!-- 解説: No.139はユーザーハイブの存在とロード状態を確認し、ロード・編集しません。 -->
| 140 | `InspectRegistryAuditHistory.vbs` | Queries recent Security event 4657 entries; requires log access and auditing. | <!-- 解説: No.140はレジストリ値変更監査イベントを読み取り、ログ設定や値は変更しません。 -->

Run scripts from an elevated console only when the requested registry location
requires it. Some diagnostics can return access-denied or unavailable results
when the account lacks permission, the key is absent, or the relevant audit
policy is not enabled. These scripts do not change those conditions.

## Additional read-only diagnostics (items 141-240)

The following grouped scripts provide safe inspection equivalents for the
requested advanced topics. They do not write registry values, change policy,
modify environment variables, load hives, or connect to remote computers.
Values that may contain credentials or identifiers are not displayed.

| No. | Script | Read-only behavior |
|---:|---|---|
| 141-160 | `141_RegistryValueAudit.vbs` | Lists value names/types and immediate subkey count for a selected local key; does not display value data. |
| 161-180 | `161_EnvironmentScopeAudit.vbs` | Compares process, user, and system environment scopes; reports PATH length, duplicate entries, and missing folders without changing them. |
| 181-205 | `181_WindowsPolicyAudit.vbs` | Reports selected Windows policy/security setting states without changing them or displaying string contents. |
| 206-240 | `206_ConfigurationInventoryAudit.vbs` | Checks selected application, runtime, browser, printer, USB, and configuration key inventories; reports counts only. |

Examples:

```text
cscript //nologo 141_RegistryValueAudit.vbs HKCU "Software"
cscript //nologo 161_EnvironmentScopeAudit.vbs
cscript //nologo 181_WindowsPolicyAudit.vbs
cscript //nologo 206_ConfigurationInventoryAudit.vbs
```

These are grouped audit examples rather than state-changing implementations.
Operations such as ACL changes, hive mounting, remote registry writes,
credential storage, policy changes, deletion, deployment, and automatic
restoration remain intentionally out of scope.
<!-- 解説: この段落は必要な権限、未設定時の結果、監査ポリシー未構成時の制限を説明しています。 -->
