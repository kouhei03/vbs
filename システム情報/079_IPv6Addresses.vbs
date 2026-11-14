Option Explicit
Dim service, adapters, adapter, address, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Description, IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
found = False
For Each adapter In adapters
    If IsArray(adapter.IPAddress) Then
        For Each address In adapter.IPAddress
            If InStr(address, ":") > 0 Then
                WScript.Echo adapter.Description & ": " & address
                found = True
            End If
        Next
    End If
Next
If Not found Then WScript.Echo "No IPv6 addresses were returned."
