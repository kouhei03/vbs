' ファイル概要: サービスの状態と開始方法をCSVに出力する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo ExportServiceInventory.vbs <new-csv-path>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: サービス一覧のCSVファイルを新規作成し、処理結果を表示します。

Option Explicit

Dim service, services, item, fileSystem, outputPath, outputFile
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
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
' [処理解説] WMIの Win32_Service を照会し、Name, DisplayName, State, StartMode を取得します。
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State, StartMode FROM Win32_Service")
On Error Resume Next
' [処理解説] 上書き引数をFalseにして、新しい出力ファイルのみを作成します。
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

' [処理解説] CSVの値を引用符で囲み、値中の引用符を二重化します。
Function Csv(value)
    If IsNull(value) Or IsEmpty(value) Then value = ""
    Csv = """" & Replace(CStr(value), """", """""") & """"
End Function
