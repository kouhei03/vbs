' IPv4/IPv6アドレスの設定とIP有効状態を照会します。バインド変更は行いません。
Option Explicit
Dim service, adapters, adapter, item, values
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Description, IPAddress, IPEnabled FROM Win32_NetworkAdapterConfiguration")
For Each adapter In adapters
    If adapter.IPEnabled Then
        values = ""
        If Not IsNull(adapter.IPAddress) Then
            For Each item In adapter.IPAddress
                If values <> "" Then values = values & ", "
                values = values & item
            Next
        End If
        WScript.Echo adapter.Description & vbTab & values
    End If
Next
