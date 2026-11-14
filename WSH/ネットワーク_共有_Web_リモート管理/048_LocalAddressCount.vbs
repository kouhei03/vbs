' ローカルNICに割り当てられているIPv4アドレス数を集計します（ネットワーク走査なし）。
Option Explicit
Dim service, adapters, adapter, address, count
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Description, IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
count = 0
For Each adapter In adapters
    If Not IsNull(adapter.IPAddress) Then
        For Each address In adapter.IPAddress
            If InStr(address, ":") = 0 And Left(address, 4) <> "169." Then
                count = count + 1
                WScript.Echo address & vbTab & adapter.Description
            End If
        Next
    End If
Next
WScript.Echo "LocalIPv4AddressCount=" & count & " (not a subnet host count)."
