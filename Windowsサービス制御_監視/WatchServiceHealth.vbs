' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、services、item、serviceName、intervalText、durationText）。主な用途: WMIサービス接続、サービス一覧、現在処理中のWMI項目、対象サービス名。
Dim service, services, item, serviceName, intervalText, durationText
' 解説: このスクリプトで使う変数を宣言します（intervalSeconds、durationSeconds、startTime、found、currentState）。主な用途: 監視間隔秒数、監視継続秒数、監視開始時刻、対象が見つかったかを示すフラグ。
Dim intervalSeconds, durationSeconds, startTime, found, currentState
' 解説: 条件「WScript.Arguments.Count <> 3 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 3 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo WatchServiceHealth.vbs <service-name> <interval-seconds 5-300> <duration-seconds 5-3600>"
    WScript.Echo "Usage: cscript //nologo WatchServiceHealth.vbs <service-name> <interval-seconds 5-300> <duration-seconds 5-3600>"
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: serviceName = Trim(WScript.Arguments(0))
serviceName = Trim(WScript.Arguments(0))
' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: intervalText = Trim(WScript.Arguments(1))
intervalText = Trim(WScript.Arguments(1))
' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: durationText = Trim(WScript.Arguments(2))
durationText = Trim(WScript.Arguments(2))
' 解説: 条件「Len(serviceName) = 0 Or Not ParseWholeNumber(intervalText, 5, 300, intervalSeconds) Or Not ParseWholeNumber(durationText, 5, 3600, durationSeconds) Then」を評価し、成立した場合に続く分岐処理を実行します。
If Len(serviceName) = 0 Or _
    Not ParseWholeNumber(intervalText, 5, 300, intervalSeconds) Or _
    Not ParseWholeNumber(durationText, 5, 3600, durationSeconds) Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Invalid service name, interval, or duration."
    WScript.Echo "Invalid service name, interval, or duration."
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: startTime = Now
startTime = Now
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: currentState = ""
currentState = ""
' 解説: 条件を確認しながら繰り返す監視/待機ループを開始します。
Do
    ' 解説: WMIのWindowsサービス（Win32_Service）を読み取り専用で照会し、Name, Stateを取得します。
    Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = False
    found = False
    ' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each item In services
    For Each item In services
        ' 解説: 条件「StrComp(item.Name, serviceName, vbTextCompare) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
        If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
            ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
            found = True
            ' 解説: 条件「StrComp(item.State, currentState, vbTextCompare) <> 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
            If StrComp(item.State, currentState, vbTextCompare) <> 0 Then
                ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: currentState = item.State
                currentState = item.State
                ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: CStr(Now) & vbTab & serviceName & vbTab & currentState
                WScript.Echo CStr(Now) & vbTab & serviceName & vbTab & currentState
                ' 解説: 条件「StrComp(currentState, "Running", vbTextCompare) <> 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
                If StrComp(currentState, "Running", vbTextCompare) <> 0 Then
                    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "ALERT: service is not running; no action was taken."
                    WScript.Echo "ALERT: service is not running; no action was taken."
                ' 解説: 条件分岐ブロックを終了します。
                End If
            ' 解説: 条件分岐ブロックを終了します。
            End If
            ' 解説: 現在のForループを直ちに抜け、後続処理へ進みます。
            Exit For
        ' 解説: 条件分岐ブロックを終了します。
        End If
    ' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
    Next
    ' 解説: 条件「Not found Then」を評価し、成立した場合に続く分岐処理を実行します。
    If Not found Then
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Service not found: " & serviceName
        WScript.Echo "Service not found: " & serviceName
        ' 解説: スクリプトを終了し、呼び出し元に終了コード 3（対象未検出）を返します。
        WScript.Quit 3
    ' 解説: 条件分岐ブロックを終了します。
    End If
    ' 解説: 条件「DateDiff("s", startTime, Now) >= durationSeconds Then Exit Do」を評価し、成立した場合に続く分岐処理を実行します。
    If DateDiff("s", startTime, Now) >= durationSeconds Then Exit Do
    ' 解説: 指定したミリ秒だけ待機し、ポーリング間隔を空けてCPU負荷を抑えます: WScript.Sleep intervalSeconds * 1000
    WScript.Sleep intervalSeconds * 1000
' 解説: ループ本体を終え、継続条件に従って次の反復を行います。
Loop
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Monitoring complete."
WScript.Echo "Monitoring complete."

' 解説: 再利用可能な関数を定義します。関数名と引数: ParseWholeNumber(text, minimum, maximum, value)
Function ParseWholeNumber(text, minimum, maximum, value)
    ' 解説: このスクリプトで使う変数を宣言します（i、character）。
    Dim i, character
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: ParseWholeNumber = False
    ParseWholeNumber = False
    ' 解説: 条件「Len(text) = 0 Or Len(text) > 4 Then Exit Function」を評価し、成立した場合に続く分岐処理を実行します。
    If Len(text) = 0 Or Len(text) > 4 Then Exit Function
    ' 解説: 指定範囲の繰り返し処理を開始します: For i = 1 To Len(text)
    For i = 1 To Len(text)
        ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: character = Mid(text, i, 1)
        character = Mid(text, i, 1)
        ' 解説: 条件「InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function」を評価し、成立した場合に続く分岐処理を実行します。
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    ' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
    Next
    ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: value = CLng(text)
    value = CLng(text)
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: ParseWholeNumber = (value >= minimum And value <= maximum)
    ParseWholeNumber = (value >= minimum And value <= maximum)
' 解説: 関数定義を終了します。
End Function
