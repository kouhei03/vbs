' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、services、item、serviceName、found）。主な用途: WMIサービス接続、サービス一覧、現在処理中のWMI項目、対象サービス名、対象が見つかったかを示すフラグ。
Dim service, services, item, serviceName, found
' 解説: 条件「WScript.Arguments.Count <> 1 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo GetServiceExitCodes.vbs <service-name>"
    WScript.Echo "Usage: cscript //nologo GetServiceExitCodes.vbs <service-name>"
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: serviceName = Trim(WScript.Arguments(0))
serviceName = Trim(WScript.Arguments(0))
' 解説: 条件「Len(serviceName) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If Len(serviceName) = 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Service name must not be empty."
    WScript.Echo "Service name must not be empty."
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのWindowsサービス（Win32_Service）を読み取り専用で照会し、Name, State, ExitCode, ServiceSpecificExitCodeを取得します。
Set services = service.ExecQuery( _
    "SELECT Name, State, ExitCode, ServiceSpecificExitCode FROM Win32_Service")
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = False
found = False
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each item In services
For Each item In services
    ' 解説: 条件「StrComp(item.Name, serviceName, vbTextCompare) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
        found = True
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "State: " & item.State
        WScript.Echo "State: " & item.State
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Exit code: " & item.ExitCode
        WScript.Echo "Exit code: " & item.ExitCode
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Service-specific exit code: " & item.ServiceSpecificExitCode
        WScript.Echo "Service-specific exit code: " & _
            item.ServiceSpecificExitCode
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
