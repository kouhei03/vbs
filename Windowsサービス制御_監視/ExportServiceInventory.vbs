' ファイル概要: サービスの状態と開始方法をCSVに出力する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo ExportServiceInventory.vbs <new-csv-path>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: サービス一覧のCSVファイルを新規作成し、処理結果を表示します。

Option Explicit

Dim service, services, item, fileSystem, outputPath, outputFile
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo ExportServiceInventory.vbs <new-csv-path>"
    WScript.Quit 2
End If
outputPath = WScript.Arguments(0)
Set fileSystem = CreateObject("Scripting.FileSystemObject")
If fileSystem.FileExists(outputPath) Then
    WScript.Echo "Output file already exists; refusing to overwrite: " & outputPath
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State, StartMode FROM Win32_Service")
On Error Resume Next
Set outputFile = fileSystem.CreateTextFile(outputPath, False, True)
If Err.Number <> 0 Then
    WScript.Echo "Could not create CSV file: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

outputFile.WriteLine """Name"",""DisplayName"",""State"",""StartMode"""
For Each item In services
    outputFile.WriteLine Csv(item.Name) & "," & Csv(item.DisplayName) & "," & _
        Csv(item.State) & "," & Csv(item.StartMode)
Next
outputFile.Close
WScript.Echo "Service inventory written to: " & outputPath

Function Csv(value)
    If IsNull(value) Or IsEmpty(value) Then value = ""
    Csv = """" & Replace(CStr(value), """", """""") & """"
End Function
