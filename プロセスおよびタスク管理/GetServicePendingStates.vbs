' ファイル概要: 処理待ち状態のサービスを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetServicePendingStates.vbs
' 実行結果: 処理待ち状態のサービスを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, services, item, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State, CheckPoint, WaitHint " & _
    "FROM Win32_Service")

found = False
For Each item In services
    If StrComp(item.State, "Start Pending", vbTextCompare) = 0 Or _
        StrComp(item.State, "Stop Pending", vbTextCompare) = 0 Or _
        StrComp(item.State, "Continue Pending", vbTextCompare) = 0 Or _
        StrComp(item.State, "Pause Pending", vbTextCompare) = 0 Then
        found = True
        WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & _
            item.State & vbTab & "Checkpoint=" & item.CheckPoint & _
            vbTab & "WaitHint(ms)=" & item.WaitHint
    End If
Next

If Not found Then WScript.Echo "No services are in a pending state."
