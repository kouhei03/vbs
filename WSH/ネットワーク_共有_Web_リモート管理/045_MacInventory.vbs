' 概要: ローカル有効NICのMACアドレスと説明を新規ローカルCSVへ出力します。
' 実行方法: Usage: cscript //nologo 045_MacInventory.vbs <new-csv-file>
' 処理結果: 有効NICのMACアドレスを新規CSVへ出力。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim fso, path, output, service, adapters, adapter
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 045_MacInventory.vbs <new-csv-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If fso.FileExists(path) Then WScript.Echo "Refusing to overwrite: " & path: WScript.Quit 2
Set output = fso.CreateTextFile(path, False, True)
output.WriteLine "Description,MACAddress"
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set adapters = service.ExecQuery("SELECT Description, MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
For Each adapter In adapters
    output.WriteLine Csv(adapter.Description) & "," & Csv(adapter.MACAddress)
Next
output.Close
WScript.Echo "Inventory written: " & path
Function Csv(value)
    If IsNull(value) Then value = ""
    Csv = """" & Replace(CStr(value), """", """""") & """"
End Function
