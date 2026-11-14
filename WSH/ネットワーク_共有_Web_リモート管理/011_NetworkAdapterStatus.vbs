' 概要: ローカルネットワークアダプターの接続状態を表示します。有効・無効化は行いません。
' 実行方法: cscript //nologo 011_NetworkAdapterStatus.vbs
' 処理結果: ローカルNIC状態照会。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, adapters, adapter
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set adapters = service.ExecQuery("SELECT Description, NetConnectionID, NetConnectionStatus, MACAddress FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True")
For Each adapter In adapters
    WScript.Echo adapter.Description & vbTab & adapter.NetConnectionID & vbTab & _
        "Status=" & adapter.NetConnectionStatus & vbTab & "MAC=" & adapter.MACAddress
Next
