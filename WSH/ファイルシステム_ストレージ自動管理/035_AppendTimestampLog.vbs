' 概要: 日時付きで一行をログに追記します。改行はログ行に混入しないよう空白化します。
' 実行方法: Usage: cscript //nologo 035_AppendTimestampLog.vbs <log-file> <message>
' 処理結果: 日時付きメッセージのログ追記。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path, message, file
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 035_AppendTimestampLog.vbs <log-file> <message>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
message = Replace(Replace(WScript.Arguments(1), vbCr, " "), vbLf, " ")
' 主処理: 入力ファイルを開きます。
Set file = fso.OpenTextFile(path, 8, True, True)
file.WriteLine Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & Right("0" & Day(Now), 2) & " " & _
    Right("0" & Hour(Now), 2) & ":" & Right("0" & Minute(Now), 2) & ":" & Right("0" & Second(Now), 2) & vbTab & message
file.Close
WScript.Echo "Appended to " & path
