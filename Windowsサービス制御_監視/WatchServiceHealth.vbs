' ファイル概要: 指定サービスの状態を指定時間にわたり監視する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo WatchServiceHealth.vbs <service-name> <interval-seconds 5-300> <duration-seconds 5-3600>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定サービスの状態を指定時間にわたり監視する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, services, item, serviceName, intervalText, durationText
Dim intervalSeconds, durationSeconds, startTime, found, currentState
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo WatchServiceHealth.vbs <service-name> <interval-seconds 5-300> <duration-seconds 5-3600>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
intervalText = Trim(WScript.Arguments(1))
durationText = Trim(WScript.Arguments(2))
If Len(serviceName) = 0 Or _
    Not ParseWholeNumber(intervalText, 5, 300, intervalSeconds) Or _
    Not ParseWholeNumber(durationText, 5, 3600, durationSeconds) Then
    WScript.Echo "Invalid service name, interval, or duration."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
startTime = Now
currentState = ""
Do
    ' [処理解説] WMIの Win32_Service を照会し、Name, State を取得します。
    Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
    found = False
    For Each item In services
        If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
            found = True
            If StrComp(item.State, currentState, vbTextCompare) <> 0 Then
                currentState = item.State
                WScript.Echo CStr(Now) & vbTab & serviceName & vbTab & currentState
                If StrComp(currentState, "Running", vbTextCompare) <> 0 Then
                    WScript.Echo "ALERT: service is not running; no action was taken."
                End If
            End If
            Exit For
        End If
    Next
    If Not found Then
        WScript.Echo "Service not found: " & serviceName
        WScript.Quit 3
    End If
    If DateDiff("s", startTime, Now) >= durationSeconds Then Exit Do
    ' [処理解説] 次の処理まで指定されたミリ秒だけ待機します。
    WScript.Sleep intervalSeconds * 1000
Loop
WScript.Echo "Monitoring complete."

' [処理解説] 数字以外の入力を除外してから整数へ変換し、指定範囲に収まるか判定します。
Function ParseWholeNumber(text, minimum, maximum, value)
    Dim i, character
    ParseWholeNumber = False
    If Len(text) = 0 Or Len(text) > 4 Then Exit Function
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParseWholeNumber = (value >= minimum And value <= maximum)
End Function
