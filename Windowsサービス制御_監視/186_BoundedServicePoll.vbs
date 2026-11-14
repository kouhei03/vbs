' ファイル概要: 期限付きサービス状態ポーリング。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 186_BoundedServicePoll.vbs <service-name> <duration-1-300> <interval-1-30>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 期限付きサービス状態ポーリング処理の結果やエラーをコンソールに表示します。

' サービス状態を最大300秒、指定間隔で読み取ります。復旧操作は行いません。
Option Explicit
Dim serviceName, duration, intervalSeconds, validator, service, services, item, startTime, found
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 186_BoundedServicePoll.vbs <service-name> <duration-1-300> <interval-1-30>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(serviceName) Or Not IsNumeric(WScript.Arguments(1)) Or Not IsNumeric(WScript.Arguments(2)) Then
    WScript.Echo "Invalid arguments."
    WScript.Quit 2
End If
duration = CLng(WScript.Arguments(1))
intervalSeconds = CLng(WScript.Arguments(2))
If duration < 1 Or duration > 300 Or intervalSeconds < 1 Or intervalSeconds > 30 Then
    WScript.Echo "Duration or interval is outside the permitted range."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
startTime = Now
Do
    Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
    found = False
    For Each item In services
        If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
            WScript.Echo CStr(Now) & vbTab & item.Name & vbTab & item.State
            found = True
            Exit For
        End If
    Next
    If Not found Then
        WScript.Echo "Service not found: " & serviceName
        WScript.Quit 3
    End If
    If DateDiff("s", startTime, Now) >= duration Then Exit Do
    WScript.Sleep intervalSeconds * 1000
Loop
WScript.Echo "Monitoring finished; no service action was taken."
