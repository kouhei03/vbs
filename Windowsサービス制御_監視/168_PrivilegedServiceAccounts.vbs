' ファイル概要: 組み込みサービスアカウント使用サービスの確認。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 168_PrivilegedServiceAccounts.vbs
' 実行結果: 組み込みサービスアカウント使用サービスの確認処理の結果やエラーをコンソールに表示します。

' 組み込みサービスアカウントで実行されるサービス名を照会します。設定変更はしません。
Option Explicit
Dim service, services, item, accountName
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery("SELECT Name, StartName, State FROM Win32_Service")
For Each item In services
    If Not IsNull(item.StartName) Then
        accountName = LCase(CStr(item.StartName))
        If accountName = "localsystem" Or accountName = "nt authority\system" Or accountName = "nt authority\localservice" Or accountName = "nt authority\networkservice" Or accountName = "localservice" Or accountName = "networkservice" Then
            WScript.Echo item.Name & vbTab & item.StartName & vbTab & item.State
        End If
    End If
Next
