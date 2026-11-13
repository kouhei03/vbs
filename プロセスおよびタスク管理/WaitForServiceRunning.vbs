' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、services、item、serviceName、timeoutText、timeoutSeconds）。主な用途: WMIサービス接続、サービス一覧、現在処理中のWMI項目、対象サービス名、待機タイムアウト秒数。
Dim service, services, item, serviceName, timeoutText, timeoutSeconds
' 解説: このスクリプトで使う変数を宣言します（found、isRunning、startTime、index、character）。主な用途: 対象が見つかったかを示すフラグ、監視開始時刻。
Dim found, isRunning, startTime, index, character

' 解説: 条件「WScript.Arguments.Count <> 2 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 2 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo WaitForServiceRunning.vbs <service-name> <timeout-seconds>"
    WScript.Echo "Usage: cscript //nologo WaitForServiceRunning.vbs <service-name> <timeout-seconds>"
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: serviceName = Trim(WScript.Arguments(0))
serviceName = Trim(WScript.Arguments(0))
' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: timeoutText = Trim(WScript.Arguments(1))
timeoutText = Trim(WScript.Arguments(1))
' 解説: 条件「Len(serviceName) = 0 Or Len(timeoutText) = 0 Or Len(timeoutText) > 3 Then」を評価し、成立した場合に続く分岐処理を実行します。
If Len(serviceName) = 0 Or Len(timeoutText) = 0 Or Len(timeoutText) > 3 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "A service name and a timeout from 1 to 300 are required."
    WScript.Echo "A service name and a timeout from 1 to 300 are required."
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: 指定範囲の繰り返し処理を開始します: For index = 1 To Len(timeoutText)
For index = 1 To Len(timeoutText)
    ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: character = Mid(timeoutText, index, 1)
    character = Mid(timeoutText, index, 1)
    ' 解説: 条件「InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
    If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Timeout must be a whole number from 1 to 300."
        WScript.Echo "Timeout must be a whole number from 1 to 300."
        ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
        WScript.Quit 2
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next

' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: timeoutSeconds = CLng(timeoutText)
timeoutSeconds = CLng(timeoutText)
' 解説: 条件「timeoutSeconds < 1 Or timeoutSeconds > 300 Then」を評価し、成立した場合に続く分岐処理を実行します。
If timeoutSeconds < 1 Or timeoutSeconds > 300 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Timeout must be between 1 and 300 seconds."
    WScript.Echo "Timeout must be between 1 and 300 seconds."
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: startTime = Now
startTime = Now

' 解説: 条件を確認しながら繰り返す監視/待機ループを開始します。
Do
    ' 解説: WMIのWindowsサービス（Win32_Service）を読み取り専用で照会し、Name, Stateを取得します。
    Set services = service.ExecQuery( _
        "SELECT Name, State FROM Win32_Service")
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = False
    found = False
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: isRunning = False
    isRunning = False
    ' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each item In services
    For Each item In services
        ' 解説: 条件「StrComp(item.Name, serviceName, vbTextCompare) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
        If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
            ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
            found = True
            ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: isRunning = (StrComp(item.State, "Running", vbTextCompare) = 0)
            isRunning = (StrComp(item.State, "Running", vbTextCompare) = 0)
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
    ' 解説: 条件「isRunning Then」を評価し、成立した場合に続く分岐処理を実行します。
    If isRunning Then
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Service is running: " & serviceName
        WScript.Echo "Service is running: " & serviceName
        ' 解説: スクリプトを終了し、呼び出し元に終了コード 0（正常終了）を返します。
        WScript.Quit 0
    ' 解説: 条件分岐ブロックを終了します。
    End If
    ' 解説: 条件「DateDiff("s", startTime, Now) >= timeoutSeconds Then」を評価し、成立した場合に続く分岐処理を実行します。
    If DateDiff("s", startTime, Now) >= timeoutSeconds Then
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Timed out waiting for service: " & serviceName
        WScript.Echo "Timed out waiting for service: " & serviceName
        ' 解説: スクリプトを終了し、呼び出し元に終了コード 1（処理失敗）を返します。
        WScript.Quit 1
    ' 解説: 条件分岐ブロックを終了します。
    End If
    ' 解説: 指定したミリ秒だけ待機し、ポーリング間隔を空けてCPU負荷を抑えます: WScript.Sleep 1000
    WScript.Sleep 1000
' 解説: ループ本体を終え、継続条件に従って次の反復を行います。
Loop