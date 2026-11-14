' 概要: Wi-Fiアダプターの接続状態を表示します。プロファイルの読出しや変更はしません。
' 実行方法: cscript //nologo 022_WifiAdapterAudit.vbs
' 処理結果: Wi-Fiアダプター状態照会（プロファイル変更なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, adapters, adapter, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set adapters = service.ExecQuery("SELECT Description, NetConnectionID, NetConnectionStatus, MACAddress FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True")
found = False
For Each adapter In adapters
    If InStr(1, adapter.Description & " " & adapter.NetConnectionID, "wi-fi", 1) > 0 Or _
        InStr(1, adapter.Description & " " & adapter.NetConnectionID, "wireless", 1) > 0 Then
        WScript.Echo adapter.Description & vbTab & adapter.NetConnectionID & vbTab & "Status=" & adapter.NetConnectionStatus
        found = True
    End If
Next
If Not found Then WScript.Echo "No Wi-Fi adapter reported."
