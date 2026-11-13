' 解説: 入力ホスト名を検証してからnslookupを実行し、DNS応答を表示します。
Option Explicit
' 解説: このスクリプト内で使用する変数を宣言します: hostname, validator, shell, process, line
Dim hostname, validator, shell, process, line
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: WScript.Arguments.Count <> 1
If WScript.Arguments.Count <> 1 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Usage: cscript //nologo 085_DnsLookup.vbs <hostname>"
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 取得値または計算結果を変数へ格納します。
hostname = WScript.Arguments(0)
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set validator = New RegExp
' 解説: 取得値または計算結果を変数へ格納します。
validator.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]*$"
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: Not validator.Test(hostname)
If Not validator.Test(hostname) Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Invalid host name."
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 読み取り用コマンドを実行し、結果取得用オブジェクトを設定します。
Set process = shell.Exec("nslookup.exe " & hostname)
' 解説: 条件が成立している間、繰り返し処理します: Do While Not process.StdOut.AtEndOfStream
Do While Not process.StdOut.AtEndOfStream
    ' 解説: 取得値または計算結果を変数へ格納します。
    line = process.StdOut.ReadLine
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo line
' 解説: Doループの現在の反復を終了します。
Loop
' 解説: 条件が成立している間、繰り返し処理します: Do While Not process.StdErr.AtEndOfStream
Do While Not process.StdErr.AtEndOfStream
    ' 解説: 取得値または計算結果を変数へ格納します。
    line = process.StdErr.ReadLine
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "ERROR: " & line
' 解説: Doループの現在の反復を終了します。
Loop
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: process.ExitCode <> 0
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
