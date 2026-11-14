' 指定サービスの状態を、最大5回の段階的間隔で観察します。再起動はしません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: serviceName, validator, service, services, item, attempt, delaySeconds, found
Dim serviceName, validator, service, services, item, attempt, delaySeconds, found
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 187_ServiceBackoffObservation.vbs <service-name>"
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
    WScript.Echo "Invalid service name."
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定範囲内の回数だけ処理を繰り返します。
For attempt = 1 To 5
    ' 解説: サービス情報を読み取り専用で照会します。
    Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    found = False
    ' 解説: コレクション内の各サービスまたは項目を順番に処理します。
    For Each item In services
        ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
        If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
            ' 解説: 取得結果またはエラーをコンソールへ表示します。
            WScript.Echo "Sample " & attempt & vbTab & item.State
            ' 解説: 入力値、取得値、または計算結果を変数に格納します。
            found = True
            ' 解説: ローカルの読み取り専用診断処理を進めます。
            Exit For
        ' 解説: 条件分岐を終了します。
        End If
    ' 解説: 次の項目または反復へ進みます。
    Next
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If Not found Then
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo "Service not found."
        ' 解説: スクリプトを終了し、終了コードを返します。
        WScript.Quit 3
    ' 解説: 条件分岐を終了します。
    End If
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    delaySeconds = attempt * 2
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If attempt < 5 Then WScript.Sleep delaySeconds * 1000
' 解説: 次の項目または反復へ進みます。
Next
' 解説: 取得結果またはエラーをコンソールへ表示します。
WScript.Echo "Observation complete; no recovery action was taken."
