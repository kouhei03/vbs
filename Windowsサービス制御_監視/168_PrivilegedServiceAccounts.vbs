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
