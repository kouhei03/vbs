' ファイル概要: VPNアダプター状態照会のみ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 124_VpnAdapterStatus.vbs
' 実行結果: VPNアダプター状態照会のみ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, adapters, adapter, label
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NetworkAdapter を照会し、Name,NetConnectionStatus,NetEnabled,PhysicalAdapter を取得します。
Set adapters = service.ExecQuery("SELECT Name,NetConnectionStatus,NetEnabled,PhysicalAdapter FROM Win32_NetworkAdapter")
For Each adapter In adapters
    label = LCase(CStr(adapter.Name))
    If InStr(label,"vpn")>0 Or InStr(label,"tunnel")>0 Or InStr(label,"wireguard")>0 Then
        WScript.Echo adapter.Name & " Status=" & adapter.NetConnectionStatus & " Enabled=" & adapter.NetEnabled
    End If
Next
