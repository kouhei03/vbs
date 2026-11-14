' ファイル概要: リモートデスクトップ接続ポリシーを確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectRemoteDesktopPolicy.vbs
' 実行結果: リモートデスクトップ接続ポリシーを確認する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, denyConnections
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetDWORDValue(&H80000002, "SYSTEM\CurrentControlSet\Control\Terminal Server", "fDenyTSConnections", denyConnections)
If result <> 0 Then
    WScript.Echo "RDP policy value unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "fDenyTSConnections=" & denyConnections
