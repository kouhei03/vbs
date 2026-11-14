' ファイル概要: 指定プロセスの製品/ファイルバージョン。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 080_ProcessFileVersion.vbs <process.exe>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定プロセスの製品/ファイルバージョン処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, processName, validator, records, process, fieldName, value, fso
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 080_ProcessFileVersion.vbs <process.exe>"
    WScript.Quit 2
End If
processName = WScript.Arguments(0)
Set validator = New RegExp
' [処理解説] 引数に許可する文字を限定して名前を検証します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(processName) Then
    WScript.Echo "Use a simple executable name such as notepad.exe."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Process を照会し、Name, ProcessId, ExecutablePath を取得します。
Set records = service.ExecQuery("SELECT Name, ProcessId, ExecutablePath FROM Win32_Process WHERE Name = '" & processName & "'")
If records.Count = 0 Then
    WScript.Echo "Process not found: " & processName
    WScript.Quit 1
End If
For Each process In records
    WScript.Echo "Name=" & process.Name & " PID=" & process.ProcessId
    Set fso = CreateObject("Scripting.FileSystemObject")
    If IsNull(process.ExecutablePath) Or Len(process.ExecutablePath) = 0 Then
        WScript.Echo "Executable path unavailable."
    Else
        WScript.Echo "Path=" & process.ExecutablePath
        WScript.Echo "FileVersion=" & fso.GetFileVersion(process.ExecutablePath)
    End If
Next
