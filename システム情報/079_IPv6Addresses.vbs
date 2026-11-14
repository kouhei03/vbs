' ファイル概要: IPv6アドレス割り当て。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 079_IPv6Addresses.vbs
' 実行結果: IPv6アドレス割り当て処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, adapters, adapter, address, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NetworkAdapterConfiguration を照会し、Description, IPAddress を取得します。
Set adapters = service.ExecQuery("SELECT Description, IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
found = False
For Each adapter In adapters
    If IsArray(adapter.IPAddress) Then
        For Each address In adapter.IPAddress
            If InStr(address, ":") > 0 Then
                WScript.Echo adapter.Description & ": " & address
                found = True
            End If
        Next
    End If
Next
If Not found Then WScript.Echo "No IPv6 addresses were returned."
