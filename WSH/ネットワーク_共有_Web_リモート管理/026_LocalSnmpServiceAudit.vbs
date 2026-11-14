' ローカルSNMPサービスの存在と状態だけを照会します。ネットワーク機器へは接続しません。
Option Explicit
Dim service, records, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Name, State, StartMode FROM Win32_Service WHERE Name='SNMP'")
If records.Count = 0 Then
    WScript.Echo "Local SNMP service is not installed."
Else
    For Each item In records
        WScript.Echo item.Name & vbTab & item.State & vbTab & item.StartMode
    Next
End If
