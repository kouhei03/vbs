' ローカルの有効なネットワークインターフェースのIPv4設定を表示します。
Option Explicit
Dim service, adapters, adapter, addresses, masks, gateways, i, gatewayText
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Description, IPAddress, IPSubnet, DefaultIPGateway, MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
For Each adapter In adapters
    addresses = SafeArrayText(adapter.IPAddress)
    masks = SafeArrayText(adapter.IPSubnet)
    gateways = SafeArrayText(adapter.DefaultIPGateway)
    WScript.Echo "Adapter=" & adapter.Description
    WScript.Echo "IPv4/IPv6=" & addresses
    WScript.Echo "Subnet=" & masks
    WScript.Echo "Gateway=" & gateways
Next
Function SafeArrayText(values)
    Dim item, result
    result = ""
    If IsNull(values) Then
        SafeArrayText = "(none)"
        Exit Function
    End If
    For Each item In values
        If result <> "" Then result = result & ", "
        result = result & item
    Next
    SafeArrayText = result
End Function
