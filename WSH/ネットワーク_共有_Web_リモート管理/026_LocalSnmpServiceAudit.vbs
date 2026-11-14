' 概要: ローカルSNMPサービスの存在と状態だけを照会します。ネットワーク機器へは接続しません。
' 実行方法: cscript //nologo 026_LocalSnmpServiceAudit.vbs
' 処理結果: ローカルSNMPサービス状態照会。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, records, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT Name, State, StartMode FROM Win32_Service WHERE Name='SNMP'")
If records.Count = 0 Then
    WScript.Echo "Local SNMP service is not installed."
Else
    For Each item In records
        WScript.Echo item.Name & vbTab & item.State & vbTab & item.StartMode
    Next
End If
