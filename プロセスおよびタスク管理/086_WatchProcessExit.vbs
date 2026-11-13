' 解説: 指定プロセスが実行中かを一定時間だけ監視し、終了を報告します。自動再起動は行いません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, records, processName, validator, timeoutSeconds, wasRunning, elapsed
Dim service, records, processName, validator, timeoutSeconds, wasRunning, elapsed
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count <> 2
If WScript.Arguments.Count <> 2 Then WScript.Echo "Usage: cscript //nologo 086_WatchProcessExit.vbs <process.exe> <timeout-seconds-1-300>": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
processName = WScript.Arguments(0)
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set validator = New RegExp
' 解説: 取得値または計算結果を変数に格納します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
' 解説: 条件を判定し、必要な分岐を選択します: If Not validator.Test(processName) Or Not IsNumeric(WScript.Arguments(1))
If Not validator.Test(processName) Or Not IsNumeric(WScript.Arguments(1)) Then WScript.Echo "Invalid arguments.": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
timeoutSeconds = CLng(WScript.Arguments(1))
' 解説: 条件を判定し、必要な分岐を選択します: If timeoutSeconds < 1 Or timeoutSeconds > 300
If timeoutSeconds < 1 Or timeoutSeconds > 300 Then WScript.Echo "Timeout must be 1..300 seconds.": WScript.Quit 2
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set records = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
' 解説: 取得値または計算結果を変数に格納します。
wasRunning = (records.Count > 0)
' 解説: 条件を判定し、必要な分岐を選択します: If Not wasRunning
If Not wasRunning Then WScript.Echo "Process is not running: " & processName: WScript.Quit 1
' 解説: 指定範囲の要素を順番に処理します: For elapsed = 1 To timeoutSeconds
For elapsed = 1 To timeoutSeconds
    ' 解説: この文でローカル診断処理を進めます: WScript.Sleep 1000
    WScript.Sleep 1000
    ' 解説: 指定条件に合う情報を読み取り専用で照会します。
    Set records = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
    ' 解説: 条件を判定し、必要な分岐を選択します: If records.Count = 0
    If records.Count = 0 Then WScript.Echo "Process exit observed after " & elapsed & " seconds.": WScript.Quit 0
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Process remained present for the observation window."
