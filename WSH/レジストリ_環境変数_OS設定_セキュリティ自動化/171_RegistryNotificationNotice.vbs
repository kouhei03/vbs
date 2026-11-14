' 概要: レジストリ変更通知は送信しません。設定を変更していないことを明示する案内です。
' 実行方法: cscript //nologo 171_RegistryNotificationNotice.vbs
' 処理結果: 設定変更・通知を行わない案内。処理結果や状態はコンソールに表示されます。
' コード解説: Windowsの状態または固定データを確認し、結果を表示します。
Option Explicit
' 取得した値または処理状態をコンソールへ表示します。
' 主処理: 設定変更・通知を行わない案内。
WScript.Echo "No registry change was made; no WM_SETTINGCHANGE notification was sent."
WScript.Echo "Only notify applications after an intentional, approved setting change."
