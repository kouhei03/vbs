' 概要: 指定した単一UNC共有へのアクセスを確認し、読み取り専用の診断項目を表示します。
' 実行方法: Usage: cscript //nologo 036_ShareAccessChecklist.vbs <UNC-share-path>
' 処理結果: 単一UNC共有へのアクセス確認と診断項目表示。処理結果や状態はコンソールに表示されます。
' コード解説: 外部コマンドの標準出力・標準エラーを読み取って結果を処理します。
Option Explicit
Dim path, re, fso, shell, runner
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 036_ShareAccessChecklist.vbs <UNC-share-path>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^\\\\[A-Za-z0-9.-]+\\[A-Za-z0-9_$.-]+(?:\\[A-Za-z0-9 _.$-]+)*$"
If Not re.Test(path) Then
    WScript.Echo "Provide one simple UNC share path."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
WScript.Echo "SharePath=" & path
If fso.FolderExists(path) Then
    WScript.Echo "Share access: available"
Else
    WScript.Echo "Share access: unavailable or access denied"
End If
Set shell = CreateObject("WScript.Shell")
' 主処理: 外部コマンドの出力を読み取ります。
Set runner = shell.Exec("net.exe use")
Do While Not runner.StdOut.AtEndOfStream
    WScript.Echo runner.StdOut.ReadLine
Loop
WScript.Echo "Also check name resolution, VPN state, server availability, and assigned permissions."
