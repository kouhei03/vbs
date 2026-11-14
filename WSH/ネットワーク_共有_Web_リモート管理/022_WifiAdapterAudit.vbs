' Wi-Fiアダプターの接続状態を表示します。プロファイルの読出しや変更はしません。
Option Explicit
Dim service, adapters, adapter, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Description, NetConnectionID, NetConnectionStatus, MACAddress FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True")
found = False
For Each adapter In adapters
    If InStr(1, adapter.Description & " " & adapter.NetConnectionID, "wi-fi", 1) > 0 Or _
        InStr(1, adapter.Description & " " & adapter.NetConnectionID, "wireless", 1) > 0 Then
        WScript.Echo adapter.Description & vbTab & adapter.NetConnectionID & vbTab & "Status=" & adapter.NetConnectionStatus
        found = True
    End If
Next
If Not found Then WScript.Echo "No Wi-Fi adapter reported."
