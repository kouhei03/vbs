' サービス失敗後の再起動遅延設定を照会します。再起動はしません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: serviceName, validator
Dim serviceName, validator
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 114_ServiceFailureRestartDelay.vbs <service-name>"
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: 入力値、取得値、または計算結果を変数に格納します。
serviceName = Trim(WScript.Arguments(0))
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set validator = New RegExp
' 解説: 入力値、取得値、または計算結果を変数に格納します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Not validator.Test(serviceName) Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Use a service name containing only letters, digits, dot, underscore, or hyphen."
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: このスクリプトで使用する変数を宣言します: shell, process, line, commandLine
Dim shell, process, line, commandLine
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 入力値、取得値、または計算結果を変数に格納します。
commandLine = "sc.exe qfailure " & serviceName
' 解説: サービス構成を表示する固定の読み取り専用コマンドを実行します。
Set process = shell.Exec(commandLine)
' 解説: 条件に従って処理を繰り返します。
Do While Not process.StdOut.AtEndOfStream
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    line = process.StdOut.ReadLine
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo line
' 解説: ループ本体を終え、次の反復条件を確認します。
Loop
' 解説: 条件に従って処理を繰り返します。
Do While Not process.StdErr.AtEndOfStream
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    line = process.StdErr.ReadLine
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "ERROR: " & line
' 解説: ループ本体を終え、次の反復条件を確認します。
Loop
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
