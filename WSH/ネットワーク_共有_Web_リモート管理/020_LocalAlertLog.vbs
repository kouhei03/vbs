' 概要: SMTP送信の代わりにローカルログへ日時付きアラートを書き込みます。
' 実行方法: Usage: cscript //nologo 020_LocalAlertLog.vbs <log-file> <message>
' 処理結果: SMTP送信の代替となるローカルアラート記録。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, path, message, output
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 020_LocalAlertLog.vbs <log-file> <message>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
message = Replace(Replace(WScript.Arguments(1), vbCr, " "), vbLf, " ")
' 主処理: 入力ファイルを開きます。
Set output = fso.OpenTextFile(path, 8, True, True)
output.WriteLine Now & vbTab & "ALERT" & vbTab & message
output.Close
WScript.Echo "Local alert recorded: " & path
