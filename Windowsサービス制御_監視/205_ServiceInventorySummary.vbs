' ローカルサービスの状態・スタートアップ種別の件数を集計します。構成変更はしません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: service, services, item, runningCount, stoppedCount, otherCount, autoCount, manualCount, disabledCount
Dim service, services, item, runningCount, stoppedCount, otherCount, autoCount, manualCount, disabledCount
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: サービス情報を読み取り専用で照会します。
Set services = service.ExecQuery("SELECT State, StartMode FROM Win32_Service")
' 解説: 入力値、取得値、または計算結果を変数に格納します。
runningCount = 0
' 解説: 入力値、取得値、または計算結果を変数に格納します。
stoppedCount = 0
' 解説: 入力値、取得値、または計算結果を変数に格納します。
otherCount = 0
' 解説: 入力値、取得値、または計算結果を変数に格納します。
autoCount = 0
' 解説: 入力値、取得値、または計算結果を変数に格納します。
manualCount = 0
' 解説: 入力値、取得値、または計算結果を変数に格納します。
disabledCount = 0
' 解説: コレクション内の各サービスまたは項目を順番に処理します。
For Each item In services
    ' 解説: 値に応じた複数分岐を開始します。
    Select Case LCase(CStr(item.State))
        ' 解説: 指定値に一致した場合の処理です。
        Case "running": runningCount = runningCount + 1
        ' 解説: 指定値に一致した場合の処理です。
        Case "stopped": stoppedCount = stoppedCount + 1
        ' 解説: 指定値に一致した場合の処理です。
        Case Else: otherCount = otherCount + 1
    ' 解説: Select Case分岐を終了します。
    End Select
    ' 解説: 値に応じた複数分岐を開始します。
    Select Case LCase(CStr(item.StartMode))
        ' 解説: 指定値に一致した場合の処理です。
        Case "auto": autoCount = autoCount + 1
        ' 解説: 指定値に一致した場合の処理です。
        Case "manual": manualCount = manualCount + 1
        ' 解説: 指定値に一致した場合の処理です。
        Case "disabled": disabledCount = disabledCount + 1
    ' 解説: Select Case分岐を終了します。
    End Select
' 解説: 次の項目または反復へ進みます。
Next
' 解説: 取得結果またはエラーをコンソールへ表示します。
WScript.Echo "Running=" & runningCount & " Stopped=" & stoppedCount & " Other=" & otherCount
' 解説: 取得結果またはエラーをコンソールへ表示します。
WScript.Echo "Automatic=" & autoCount & " Manual=" & manualCount & " Disabled=" & disabledCount
