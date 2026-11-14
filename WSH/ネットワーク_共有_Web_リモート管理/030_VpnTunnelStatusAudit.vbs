' ローカルVPN関連アダプターと状態を表示します。切断・再接続はしません。
Option Explicit
Dim service, adapters, adapter, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Description, NetConnectionID, NetConnectionStatus FROM Win32_NetworkAdapter")
found = False
For Each adapter In adapters
    If InStr(1, adapter.Description & " " & adapter.NetConnectionID, "vpn", 1) > 0 Then
        WScript.Echo adapter.Description & vbTab & adapter.NetConnectionID & vbTab & "Status=" & adapter.NetConnectionStatus
        found = True
    End If
Next
If Not found Then WScript.Echo "No VPN-related adapter reported."
