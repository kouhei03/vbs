' ファイル概要: 指定サービスが実行中になるまで待機する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo WaitForServiceRunning.vbs <service-name> <timeout-seconds 1-300>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定サービスが実行中になるまで待機する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, services, item, serviceName, timeoutText, timeoutSeconds
Dim found, running, startTime, index, character
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo WaitForServiceRunning.vbs <service-name> <timeout-seconds 1-300>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
timeoutText = Trim(WScript.Arguments(1))
If Len(serviceName) = 0 Or Len(timeoutText) = 0 Or Len(timeoutText) > 3 Then
    WScript.Echo "A service name and timeout from 1 to 300 are required."
    WScript.Quit 2
End If
For index = 1 To Len(timeoutText)
    character = Mid(timeoutText, index, 1)
    If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then
        WScript.Echo "Timeout must be a whole number."
        WScript.Quit 2
    End If
Next
timeoutSeconds = CLng(timeoutText)
If timeoutSeconds < 1 Or timeoutSeconds > 300 Then
    WScript.Echo "Timeout must be between 1 and 300 seconds."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
startTime = Now
Do
    ' [処理解説] WMIの Win32_Service を照会し、Name, State を取得します。
    Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
    found = False
    running = False
    For Each item In services
        If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
            found = True
            running = (StrComp(item.State, "Running", vbTextCompare) = 0)
            Exit For
        End If
    Next
    If Not found Then
        WScript.Echo "Service not found: " & serviceName
        WScript.Quit 3
    End If
    If running Then
        WScript.Echo "Service is running: " & serviceName
        WScript.Quit 0
    End If
    If DateDiff("s", startTime, Now) >= timeoutSeconds Then
        WScript.Echo "Timed out waiting for service: " & serviceName
        WScript.Quit 1
    End If
    ' [処理解説] 次の処理まで指定されたミリ秒だけ待機します。
    WScript.Sleep 1000
Loop
