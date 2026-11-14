' NICの電源管理に関するWMIプロパティを照会します。レジストリや電源設定は変更しません。
Option Explicit
Dim service, adapters, adapter
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Name, PowerManagementSupported, PowerManagementCapabilities FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True")
For Each adapter In adapters
    WScript.Echo "Adapter=" & adapter.Name & " PowerManagementSupported=" & adapter.PowerManagementSupported
    WScript.Echo "Capabilities=" & ArrayText(adapter.PowerManagementCapabilities)
Next
Function ArrayText(values)
    Dim item, result
    result = ""
    If Not IsNull(values) Then
        For Each item In values
            If result <> "" Then result = result & ","
            result = result & item
        Next
    End If
    ArrayText = result
End Function
