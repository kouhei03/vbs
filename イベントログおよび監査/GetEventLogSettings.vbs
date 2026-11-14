' ファイル概要: イベントログの容量と上書き設定を表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetEventLogSettings.vbs
' 実行結果: イベントログの容量と上書き設定を表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, logs, logItem
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTEventLogFile を照会し、Name, FileName, FileSize, MaxFileSize, OverwriteOutdated, OverwritePolicy を取得します。
Set logs = service.ExecQuery( _
    "SELECT Name, FileName, FileSize, MaxFileSize, " & _
    "OverwriteOutdated, OverwritePolicy FROM Win32_NTEventLogFile")

WScript.Echo "Log" & vbTab & "Current size (bytes)" & vbTab & _
    "Maximum size (bytes)" & vbTab & "Overwrite policy"
For Each logItem In logs
    WScript.Echo logItem.Name & vbTab & logItem.FileSize & vbTab & _
        logItem.MaxFileSize & vbTab & logItem.OverwritePolicy
Next
