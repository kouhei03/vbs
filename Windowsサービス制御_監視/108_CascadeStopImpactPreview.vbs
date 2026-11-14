' 停止した場合に影響を受け得る依存サービスを表示します。停止は行いません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: serviceName, validator, service, relations, relation, services, item, marker, found, stateText
Dim serviceName, validator, service, relations, relation, services, item, marker, found, stateText
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 108_CascadeStopImpactPreview.vbs <service-name>"
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
' 解説: サービス情報を読み取り専用で照会します。
Set relations = service.ExecQuery("SELECT Antecedent, Dependent FROM Win32_DependentService")
' 解説: サービス情報を読み取り専用で照会します。
Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
' 解説: 入力値、取得値、または計算結果を変数に格納します。
marker = "Name=""" & serviceName & """"
' 解説: 入力値、取得値、または計算結果を変数に格納します。
found = False
' 解説: コレクション内の各サービスまたは項目を順番に処理します。
For Each relation In relations
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If InStr(1, relation.Antecedent, marker, vbTextCompare) > 0 Then
        ' 解説: 入力値、取得値、または計算結果を変数に格納します。
        stateText = "Unknown"
        ' 解説: コレクション内の各サービスまたは項目を順番に処理します。
        For Each item In services
            ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
            If InStr(1, relation.Dependent, "Name=""" & item.Name & """", vbTextCompare) > 0 Then stateText = item.State
        ' 解説: 次の項目または反復へ進みます。
        Next
        ' 解説: 取得結果またはエラーをコンソールへ表示します。
        WScript.Echo relation.Dependent & vbTab & stateText
        ' 解説: 入力値、取得値、または計算結果を変数に格納します。
        found = True
    ' 解説: 条件分岐を終了します。
    End If
' 解説: 次の項目または反復へ進みます。
Next
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Not found Then WScript.Echo "No dependent services found."
