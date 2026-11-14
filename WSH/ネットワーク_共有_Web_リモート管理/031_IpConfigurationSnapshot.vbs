' 現在のローカルIP設定を新規テキストファイルへ記録します。
Option Explicit
Dim fso, outputPath, output, service, adapters, adapter
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 031_IpConfigurationSnapshot.vbs <new-snapshot-file>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
outputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
If fso.FileExists(outputPath) Then WScript.Echo "Refusing to overwrite: " & outputPath: WScript.Quit 2
Set output = fso.CreateTextFile(outputPath, False, True)
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Description, IPAddress, IPSubnet, DefaultIPGateway FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
For Each adapter In adapters
    output.WriteLine "Adapter=" & adapter.Description
    output.WriteLine "IP=" & JoinValues(adapter.IPAddress)
    output.WriteLine "Subnet=" & JoinValues(adapter.IPSubnet)
    output.WriteLine "Gateway=" & JoinValues(adapter.DefaultIPGateway)
Next
output.Close
WScript.Echo "Snapshot written: " & outputPath
Function JoinValues(values)
    Dim item, result
    result = ""
    If Not IsNull(values) Then
        For Each item In values
            If result <> "" Then result = result & ","
            result = result & item
        Next
    End If
    JoinValues = result
End Function
