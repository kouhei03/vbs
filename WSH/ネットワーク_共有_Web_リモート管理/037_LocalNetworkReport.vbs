' 概要: 公開天気・為替サービスへ接続せず、ローカルIP設定を新規レポートに保存します。
' 実行方法: Usage: cscript //nologo 037_LocalNetworkReport.vbs <new-report-file>
' 処理結果: ローカルIP・MAC情報の新規レポート。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim fso, path, output, service, adapters, adapter, item
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 037_LocalNetworkReport.vbs <new-report-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If fso.FileExists(path) Then WScript.Echo "Refusing to overwrite: " & path: WScript.Quit 2
Set output = fso.CreateTextFile(path, False, True)
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set adapters = service.ExecQuery("SELECT Description, MACAddress, IPAddress, DefaultIPGateway FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
For Each adapter In adapters
    output.WriteLine "Adapter=" & adapter.Description
    output.WriteLine "MAC=" & ScalarText(adapter.MACAddress)
    output.WriteLine "IP=" & ValuesText(adapter.IPAddress)
    output.WriteLine "Gateway=" & ValuesText(adapter.DefaultIPGateway)
Next
output.Close
WScript.Echo "Local-only report: " & path
Function ValuesText(values)
    Dim value, result
    result = ""
    If Not IsNull(values) Then
        For Each value In values
            If result <> "" Then result = result & ","
            result = result & value
        Next
    End If
    ValuesText = result
End Function
Function ScalarText(value)
    If IsNull(value) Then
        ScalarText = ""
    Else
        ScalarText = CStr(value)
    End If
End Function
