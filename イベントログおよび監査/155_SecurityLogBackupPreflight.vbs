' ファイル概要: セキュリティログのメタデータを確認する（バックアップなし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 155_SecurityLogBackupPreflight.vbs
' 実行結果: セキュリティログのメタデータを確認する（バックアップなし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, logs, logItem, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set logs = service.ExecQuery( _
    "SELECT Name, FileName, FileSize, MaxFileSize FROM Win32_NTEventLogFile " & _
    "WHERE Name='Security'")
found = False
For Each logItem In logs
    found = True
    WScript.Echo "Security log file: " & logItem.FileName
    WScript.Echo "Current bytes: " & logItem.FileSize
    WScript.Echo "Maximum bytes: " & logItem.MaxFileSize
Next
If Not found Then
    WScript.Echo "Security log metadata is unavailable; check permissions."
End If
WScript.Echo "Preflight only; no backup or export was created."
