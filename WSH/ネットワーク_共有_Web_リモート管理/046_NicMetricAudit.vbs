' 概要: IP有効NICのIPv4/IPv6接続メトリックを表示します。優先度は変更しません。
' 実行方法: cscript //nologo 046_NicMetricAudit.vbs
' 処理結果: NICメトリック照会（変更なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, adapters, adapter
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set adapters = service.ExecQuery("SELECT Description, InterfaceIndex, IPConnectionMetric FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
For Each adapter In adapters
    WScript.Echo adapter.Description & vbTab & "InterfaceIndex=" & adapter.InterfaceIndex & _
        " Metric=" & adapter.IPConnectionMetric
Next
