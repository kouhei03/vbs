' 概要: ローカルVPN関連アダプターと状態を表示します。切断・再接続はしません。
' 実行方法: cscript //nologo 030_VpnTunnelStatusAudit.vbs
' 処理結果: VPN関連アダプター状態照会。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, adapters, adapter, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set adapters = service.ExecQuery("SELECT Description, NetConnectionID, NetConnectionStatus FROM Win32_NetworkAdapter")
found = False
For Each adapter In adapters
    If InStr(1, adapter.Description & " " & adapter.NetConnectionID, "vpn", 1) > 0 Then
        WScript.Echo adapter.Description & vbTab & adapter.NetConnectionID & vbTab & "Status=" & adapter.NetConnectionStatus
        found = True
    End If
Next
If Not found Then WScript.Echo "No VPN-related adapter reported."
