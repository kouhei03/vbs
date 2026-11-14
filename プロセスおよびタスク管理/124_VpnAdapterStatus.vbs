Option Explicit
Dim service, adapters, adapter, label
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Name,NetConnectionStatus,NetEnabled,PhysicalAdapter FROM Win32_NetworkAdapter")
For Each adapter In adapters
    label = LCase(CStr(adapter.Name))
    If InStr(label,"vpn")>0 Or InStr(label,"tunnel")>0 Or InStr(label,"wireguard")>0 Then
        WScript.Echo adapter.Name & " Status=" & adapter.NetConnectionStatus & " Enabled=" & adapter.NetEnabled
    End If
Next
