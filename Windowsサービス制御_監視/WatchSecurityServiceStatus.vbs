' ファイル概要: 主要なセキュリティサービスの状態を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo WatchSecurityServiceStatus.vbs
' 実行結果: 主要なセキュリティサービスの状態を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, services, item, watchedNames, watchedName, found, stateText
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State FROM Win32_Service")
watchedNames = Array("WinDefend", "MpsSvc", "wscsvc", "Sense")

For Each watchedName In watchedNames
    found = False
    For Each item In services
        If StrComp(item.Name, watchedName, vbTextCompare) = 0 Then
            found = True
            stateText = item.State
            WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & stateText
            Exit For
        End If
    Next
    If Not found Then
        WScript.Echo watchedName & vbTab & "Not installed or not visible"
    End If
Next
