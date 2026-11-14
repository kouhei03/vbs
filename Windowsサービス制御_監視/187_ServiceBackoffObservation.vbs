' ファイル概要: 段階的間隔での状態観察のみ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 187_ServiceBackoffObservation.vbs <service-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 段階的間隔での状態観察のみ処理の結果やエラーをコンソールに表示します。

' 指定サービスの状態を、最大5回の段階的間隔で観察します。再起動はしません。
Option Explicit
Dim serviceName, validator, service, services, item, attempt, delaySeconds, found
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 187_ServiceBackoffObservation.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
Set validator = New RegExp
' [処理解説] 引数に許可する文字を限定して名前を検証します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(serviceName) Then
    WScript.Echo "Invalid service name."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
For attempt = 1 To 5
    ' [処理解説] WMIの Win32_Service を照会し、Name, State を取得します。
    Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
    found = False
    For Each item In services
        If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
            WScript.Echo "Sample " & attempt & vbTab & item.State
            found = True
            Exit For
        End If
    Next
    If Not found Then
        WScript.Echo "Service not found."
        WScript.Quit 3
    End If
    delaySeconds = attempt * 2
    ' [処理解説] 次の処理まで指定されたミリ秒だけ待機します。
    If attempt < 5 Then WScript.Sleep delaySeconds * 1000
Next
WScript.Echo "Observation complete; no recovery action was taken."
