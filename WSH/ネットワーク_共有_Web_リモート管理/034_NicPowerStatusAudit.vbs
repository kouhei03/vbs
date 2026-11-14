' ネットワークアダプターの電源管理に関するプロパティを照会します（変更なし）。
Option Explicit
Dim service, adapters, adapter, property
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Name, PNPDeviceID, PowerManagementSupported, PowerManagementCapabilities FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True")
For Each adapter In adapters
    WScript.Echo "Adapter=" & adapter.Name & " PowerManagementSupported=" & adapter.PowerManagementSupported
    WScript.Echo "PowerManagementCapabilities=" & JoinValues(adapter.PowerManagementCapabilities)
Next
Function JoinValues(values)
    Dim item, result
    result = ""
    If Not IsNull(values) Then
        For Each item In values
            If result <> "" Then result = result & ","
            result = result & item
        Next
    End If
    JoinValues = result
End Function
