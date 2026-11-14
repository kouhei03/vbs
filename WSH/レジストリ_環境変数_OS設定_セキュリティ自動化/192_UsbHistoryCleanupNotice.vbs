' 概要: USB device historyを削除しません。監査要件・証拠保全への影響を案内します。
' 実行方法: cscript //nologo 192_UsbHistoryCleanupNotice.vbs
' 処理結果: USB履歴を削除しない案内。処理結果や状態はコンソールに表示されます。
' コード解説: Windowsの状態または固定データを確認し、結果を表示します。
Option Explicit
' 取得した値または処理状態をコンソールへ表示します。
' 主処理: USB履歴を削除しない案内。
WScript.Echo "No USBSTOR history or device record was deleted."
WScript.Echo "Preserve device history for authorized auditing and incident response."
