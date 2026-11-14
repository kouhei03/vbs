' ファイル概要: プロセス終了の期限付き監視（再起動なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 086_WatchProcessExit.vbs <process.exe> <timeout-seconds-1-300>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: プロセス終了の期限付き監視（再起動なし）処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, processName, validator, timeoutSeconds, wasRunning, elapsed
If WScript.Arguments.Count <> 2 Then WScript.Echo "Usage: cscript //nologo 086_WatchProcessExit.vbs <process.exe> <timeout-seconds-1-300>": WScript.Quit 2
processName = WScript.Arguments(0)
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(processName) Or Not IsNumeric(WScript.Arguments(1)) Then WScript.Echo "Invalid arguments.": WScript.Quit 2
timeoutSeconds = CLng(WScript.Arguments(1))
If timeoutSeconds < 1 Or timeoutSeconds > 300 Then WScript.Echo "Timeout must be 1..300 seconds.": WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
wasRunning = (records.Count > 0)
If Not wasRunning Then WScript.Echo "Process is not running: " & processName: WScript.Quit 1
For elapsed = 1 To timeoutSeconds
    WScript.Sleep 1000
    Set records = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
    If records.Count = 0 Then WScript.Echo "Process exit observed after " & elapsed & " seconds.": WScript.Quit 0
Next
WScript.Echo "Process remained present for the observation window."
