' サービス状態を最大300秒、指定間隔で読み取ります。復旧操作は行いません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: serviceName, duration, intervalSeconds, validator, service, services, item, startTime, found
Dim serviceName, duration, intervalSeconds, validator, service, services, item, startTime, found
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If WScript.Arguments.Count <> 3 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 186_BoundedServicePoll.vbs <service-name> <duration-1-300> <interval-1-30>"
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
If Not validator.Test(serviceName) Or Not IsNumeric(WScript.Arguments(1)) Or Not IsNumeric(WScript.Arguments(2)) Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Invalid arguments."
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: 入力値、取得値、または計算結果を変数に格納します。
duration = CLng(WScript.Arguments(1))
' 解説: 入力値、取得値、または計算結果を変数に格納します。
intervalSeconds = CLng(WScript.Arguments(2))
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If duration < 1 Or duration > 300 Or intervalSeconds < 1 Or intervalSeconds > 30 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Duration or interval is outside the permitted range."
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐を終了します。
End If
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 入力値、取得値、または計算結果を変数に格納します。
startTime = Now
' 解説: ローカルの読み取り専用診断処理を進めます。
Do
    ' 解説: サービス情報を読み取り専用で照会します。
    Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    found = False
    ' 解説: コレクション内の各サービスまたは項目を順番に処理します。
    For Each item In services
        ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
        If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
            ' 解説: 取得結果またはエラーをコンソールへ表示します。
            WScript.Echo CStr(Now) & vbTab & item.Name & vbTab & item.State
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
        WScript.Echo "Service not found: " & serviceName
        ' 解説: スクリプトを終了し、終了コードを返します。
        WScript.Quit 3
    ' 解説: 条件分岐を終了します。
    End If
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If DateDiff("s", startTime, Now) >= duration Then Exit Do
    ' 解説: 次の監視サンプルまで待機し、ポーリング頻度を抑えます。
    WScript.Sleep intervalSeconds * 1000
' 解説: ループ本体を終え、次の反復条件を確認します。
Loop
' 解説: 取得結果またはエラーをコンソールへ表示します。
WScript.Echo "Monitoring finished; no service action was taken."
