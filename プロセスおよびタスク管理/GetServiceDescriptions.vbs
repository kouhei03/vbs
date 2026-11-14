' ファイル概要: サービスの表示名と説明を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetServiceDescriptions.vbs
' 実行結果: サービスの表示名と説明を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, services, item, description
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, Description FROM Win32_Service")

WScript.Echo "Service name" & vbTab & "Display name" & vbTab & "Description"
For Each item In services
    description = item.Description
    If IsNull(description) Then
        description = "(no description)"
    ElseIf Len(description) = 0 Then
        description = "(no description)"
    End If
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & description
Next
