' IP有効NICのIPv4/IPv6接続メトリックを表示します。優先度は変更しません。
Option Explicit
Dim service, adapters, adapter
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Description, InterfaceIndex, IPConnectionMetric FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
For Each adapter In adapters
    WScript.Echo adapter.Description & vbTab & "InterfaceIndex=" & adapter.InterfaceIndex & _
        " Metric=" & adapter.IPConnectionMetric
Next
