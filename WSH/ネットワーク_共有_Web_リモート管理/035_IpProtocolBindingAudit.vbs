' 概要: IPv4/IPv6アドレスの設定とIP有効状態を照会します。バインド変更は行いません。
' 実行方法: cscript //nologo 035_IpProtocolBindingAudit.vbs
' 処理結果: IPv4/IPv6アドレス構成の確認。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, adapters, adapter, item, values
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set adapters = service.ExecQuery("SELECT Description, IPAddress, IPEnabled FROM Win32_NetworkAdapterConfiguration")
For Each adapter In adapters
    If adapter.IPEnabled Then
        values = ""
        If Not IsNull(adapter.IPAddress) Then
            For Each item In adapter.IPAddress
                If values <> "" Then values = values & ", "
                values = values & item
            Next
        End If
        WScript.Echo adapter.Description & vbTab & values
    End If
Next
