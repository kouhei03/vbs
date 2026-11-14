' ファイル概要: OSの診断概要を新規ローカルファイルに出力する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 170_SystemDiagnosticSummary.vbs <new-output-file>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: OS名・バージョン・ビルド・アーキテクチャ・起動時刻を新規ファイルに保存し、処理結果を表示します。

Option Explicit

Dim outputPath, fso, file, service, systems, systemItem
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 170_SystemDiagnosticSummary.vbs <new-output-file>"
    WScript.Quit 2
End If
outputPath = WScript.Arguments(0)
If Left(outputPath, 2) = "\\" Then
    WScript.Echo "UNC output paths are not allowed; choose a local destination."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(outputPath) Or fso.FolderExists(outputPath) Then
    WScript.Echo "Output path already exists; refusing to overwrite."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery( _
    "SELECT Caption, Version, BuildNumber, OSArchitecture, LastBootUpTime " & _
    "FROM Win32_OperatingSystem")
On Error Resume Next
Set file = fso.CreateTextFile(outputPath, False, True)
If Err.Number <> 0 Then
    WScript.Echo "Could not create output file: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
For Each systemItem In systems
    file.WriteLine "OS=" & systemItem.Caption
    file.WriteLine "Version=" & systemItem.Version
    file.WriteLine "Build=" & systemItem.BuildNumber
    file.WriteLine "Architecture=" & systemItem.OSArchitecture
    file.WriteLine "LastBootUpTime=" & systemItem.LastBootUpTime
Next
file.Close
WScript.Echo "Local diagnostic summary created; no event logs or credentials were exported."
