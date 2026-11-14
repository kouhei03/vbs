' 概要: ネットワークアダプターの電源管理に関するプロパティを照会します（変更なし）。
' 実行方法: cscript //nologo 034_NicPowerStatusAudit.vbs
' 処理結果: NIC電源管理プロパティ確認。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, adapters, adapter, property
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
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
