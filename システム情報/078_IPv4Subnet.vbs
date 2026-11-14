' ファイル概要: IPv4アドレスとサブネット。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 078_IPv4Subnet.vbs
' 実行結果: IPv4アドレスとサブネット処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, adapters, adapter, addresses, subnets, index, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Description, IPAddress, IPSubnet FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
found = False
For Each adapter In adapters
    addresses = adapter.IPAddress
    subnets = adapter.IPSubnet
    If IsArray(addresses) Then
        For index = LBound(addresses) To UBound(addresses)
            If InStr(addresses(index), ":") = 0 Then
                found = True
                WScript.Echo adapter.Description & ": " & addresses(index)
                If IsArray(subnets) Then
                    If index <= UBound(subnets) Then WScript.Echo "Subnet: " & subnets(index)
                End If
            End If
        Next
    End If
Next
If Not found Then WScript.Echo "No IPv4 address was returned."
