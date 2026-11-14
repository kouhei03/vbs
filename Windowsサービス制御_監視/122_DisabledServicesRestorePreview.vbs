' 無効状態のサービスを列挙し、復旧検討用の情報を表示します。設定は変更しません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: service, services, item, count
Dim service, services, item, count
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: サービス情報を読み取り専用で照会します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode FROM Win32_Service WHERE StartMode='Disabled'")
' 解説: 入力値、取得値、または計算結果を変数に格納します。
count = 0
' 解説: コレクション内の各サービスまたは項目を順番に処理します。
For Each item In services
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State & vbTab & item.StartMode
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    count = count + 1
' 解説: 次の項目または反復へ進みます。
Next
' 解説: 取得結果またはエラーをコンソールへ表示します。
WScript.Echo "Disabled services found=" & count & "; no startup settings were changed."
