' 概要: NICの電源管理に関するWMIプロパティを照会します。レジストリや電源設定は変更しません。
' 実行方法: cscript //nologo 181_NicPowerManagementAudit.vbs
' 処理結果: NIC電源管理プロパティ確認。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, adapters, adapter
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
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
