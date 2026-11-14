' 概要: ローカルNIC/IP情報とルーティングテーブルを新規ZIPに収集します。外部送信なし。
' 実行方法: Usage: cscript //nologo 050_LocalNetworkDiagnostics.vbs <new-report.zip>
' 処理結果: ローカル情報を新規ZIPに収集。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim fso, archivePath, reportPath, tempFolder, output, service, adapters, adapter
Dim shell, runner, zipFolder, sourceFolder, deadline, lastSize, stableSince
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 050_LocalNetworkDiagnostics.vbs <new-report.zip>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
archivePath = fso.GetAbsolutePathName(WScript.Arguments(0))
If LCase(fso.GetExtensionName(archivePath)) <> "zip" Or fso.FileExists(archivePath) Or _
    fso.FolderExists(archivePath) Or Not fso.FolderExists(fso.GetParentFolderName(archivePath)) Then
    WScript.Echo "Provide a new .zip file in an existing folder."
    WScript.Quit 2
End If
tempFolder = fso.BuildPath(fso.GetSpecialFolder(2).Path, "wsh-netdiag-" & fso.GetTempName)
fso.CreateFolder tempFolder
reportPath = fso.BuildPath(tempFolder, "network-diagnostics.txt")
Set output = fso.CreateTextFile(reportPath, False, True)
output.WriteLine "Local network diagnostics " & Now
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set adapters = service.ExecQuery("SELECT Description, MACAddress, IPAddress, DefaultIPGateway FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
For Each adapter In adapters
    output.WriteLine "Adapter=" & ScalarText(adapter.Description)
    output.WriteLine "MAC=" & ScalarText(adapter.MACAddress)
    output.WriteLine "IP=" & ValuesText(adapter.IPAddress)
    output.WriteLine "Gateway=" & ValuesText(adapter.DefaultIPGateway)
Next
output.WriteLine
output.WriteLine "=== Local route table ==="
Set shell = CreateObject("WScript.Shell")
Set runner = shell.Exec("route.exe print")
Do While Not runner.StdOut.AtEndOfStream
    output.WriteLine runner.StdOut.ReadLine
Loop
If runner.ExitCode <> 0 Then
    output.Close
    WScript.Echo "route.exe failed with exit code " & runner.ExitCode
    WScript.Quit runner.ExitCode
End If
output.Close
Set output = CreateObject("ADODB.Stream")
output.Type = 2
output.Charset = "windows-1252"
output.Open
output.WriteText Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, Chr(0))
output.SaveToFile archivePath, 1
output.Close
Set shell = CreateObject("Shell.Application")
Set zipFolder = shell.NameSpace(archivePath)
deadline = DateAdd("s", 10, Now)
Do While zipFolder Is Nothing And Now < deadline
    WScript.Sleep 250
    Set zipFolder = shell.NameSpace(archivePath)
Loop
If zipFolder Is Nothing Then
    WScript.Echo "Could not open ZIP archive. Temporary report remains at " & reportPath
    WScript.Quit 1
End If
Set sourceFolder = shell.NameSpace(tempFolder)
zipFolder.CopyHere sourceFolder.Items, 20
deadline = DateAdd("s", 30, Now)
lastSize = -1
stableSince = Now
Do While Now < deadline
    If zipFolder.Items.Count > 0 Then
        If fso.GetFile(archivePath).Size <> lastSize Then
            lastSize = fso.GetFile(archivePath).Size
            stableSince = Now
        ElseIf DateDiff("s", stableSince, Now) >= 3 Then
            Exit Do
        End If
    End If
    WScript.Sleep 500
Loop
If zipFolder.Items.Count = 0 Or Now >= deadline Then
    WScript.Echo "ZIP creation timed out. Temporary report remains at " & reportPath
    WScript.Quit 1
End If
fso.DeleteFolder tempFolder, True
WScript.Echo "Local-only diagnostic ZIP created: " & archivePath
WScript.Echo "Review internal network addresses before sharing the archive."
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
