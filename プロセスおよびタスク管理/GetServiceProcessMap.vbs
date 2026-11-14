' ファイル概要: サービスと対応するプロセスIDを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetServiceProcessMap.vbs
' 実行結果: サービスと対応するプロセスIDを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, services, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Service を照会し、Name, DisplayName, State, ProcessId を取得します。
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State, ProcessId FROM Win32_Service")

WScript.Echo "Service name" & vbTab & "Display name" & vbTab & _
    "State" & vbTab & "PID"
For Each item In services
    If item.ProcessId > 0 Then
        WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & _
            item.State & vbTab & item.ProcessId
    End If
Next
