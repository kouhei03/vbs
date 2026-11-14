' ローカル全サービスの状態を、新規CSVファイルへ一度だけ出力します。既存ファイルは上書きしません。
Option Explicit
Dim fso, path, file, service, services, item
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 181_ServiceStateSnapshot.vbs <new-csv-path>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(path) Then
    WScript.Echo "Refusing to overwrite: " & path
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode, StartName, ExitCode FROM Win32_Service")
Set file = fso.CreateTextFile(path, False, True)
file.WriteLine "Name,DisplayName,State,StartMode,StartName,ExitCode"
For Each item In services
    file.WriteLine Csv(item.Name) & "," & Csv(item.DisplayName) & "," & Csv(item.State) & "," & Csv(item.StartMode) & "," & Csv(item.StartName) & "," & Csv(item.ExitCode)
Next
file.Close
WScript.Echo "Snapshot written: " & path
Function Csv(value)
    If IsNull(value) Or IsEmpty(value) Then value = ""
    Csv = """" & Replace(CStr(value), """", """""") & """"
End Function
