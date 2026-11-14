' ファイル概要: サービスの終了コードを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetServiceExitCodes.vbs
' 実行結果: サービスの終了コードを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, services, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, State, ExitCode, ServiceSpecificExitCode FROM Win32_Service")

WScript.Echo "Service name" & vbTab & "State" & vbTab & _
    "Exit code" & vbTab & "Service-specific exit code"
For Each item In services
    WScript.Echo item.Name & vbTab & item.State & vbTab & _
        item.ExitCode & vbTab & item.ServiceSpecificExitCode
Next
