' 指定間隔で限定回数だけサービス数を照会し、簡単な低頻度ポーリング例を示します。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: service, services, sample
Dim service, services, sample
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定範囲内の回数だけ処理を繰り返します。
For sample = 1 To 3
    ' 解説: サービス情報を読み取り専用で照会します。
    Set services = service.ExecQuery("SELECT Name FROM Win32_Service")
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo CStr(Now) & vbTab & "Sample=" & sample & vbTab & "Services=" & services.Count
    ' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
    If sample < 3 Then WScript.Sleep 10000
' 解説: 次の項目または反復へ進みます。
Next
