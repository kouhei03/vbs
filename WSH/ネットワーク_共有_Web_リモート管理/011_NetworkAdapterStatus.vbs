' ローカルネットワークアダプターの接続状態を表示します。有効・無効化は行いません。
Option Explicit
Dim service, adapters, adapter
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Description, NetConnectionID, NetConnectionStatus, MACAddress FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True")
For Each adapter In adapters
    WScript.Echo adapter.Description & vbTab & adapter.NetConnectionID & vbTab & _
        "Status=" & adapter.NetConnectionStatus & vbTab & "MAC=" & adapter.MACAddress
Next
