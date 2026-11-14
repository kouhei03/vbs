' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: service, services, item, wantedNames, wantedName, found
Dim service, services, item, wantedNames, wantedName, found
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: サービス情報を読み取り専用で照会します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode, ExitCode FROM Win32_Service")
' 解説: 入力値、取得値、または計算結果を変数に格納します。
wantedNames = Array("HealthService", "WinCollect", "SplunkForwarder", "Elastic-Agent")
' 解説: コレクション内の各サービスまたは項目を順番に処理します。
For Each wantedName In wantedNames
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    found = False
    ' 解説: コレクション内の各サービスまたは項目を順番に処理します。
    For Each item In services
        ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
        If StrComp(item.Name, wantedName, vbTextCompare) = 0 Then
            ' 解説: 取得結果またはエラーをコンソールへ表示します。
            WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State & vbTab & item.StartMode & vbTab & item.ExitCode
            ' 解説: 入力値、取得値、または計算結果を変数に格納します。
            found = True
            ' 解説: ローカルの読み取り専用診断処理を進めます。
            Exit For
        ' 解説: 条件分岐を終了します。
        End If
    ' 解説: 次の項目または反復へ進みます。
    Next
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If Not found Then WScript.Echo wantedName & vbTab & "Not installed or not visible"
' 解説: 次の項目または反復へ進みます。
Next
