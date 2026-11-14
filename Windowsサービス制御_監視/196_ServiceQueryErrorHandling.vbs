' WMI照会の失敗を明示的に報告する読み取り専用エラーハンドリング例です。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: service, services, errorNumber, errorDescription
Dim service, services, errorNumber, errorDescription
' 解説: この限定範囲のCOMエラーを検査できるようにします。
On Error Resume Next
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Err.Number <> 0 Then
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    errorNumber = Err.Number
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    errorDescription = Err.Description
    ' 解説: この限定範囲のCOMエラーを検査できるようにします。
    On Error GoTo 0
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "WMI connection failed: " & errorNumber & " " & errorDescription
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐を終了します。
End If
' 解説: サービス情報を読み取り専用で照会します。
Set services = service.ExecQuery("SELECT Name FROM Win32_Service")
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Err.Number <> 0 Then
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    errorNumber = Err.Number
    ' 解説: 入力値、取得値、または計算結果を変数に格納します。
    errorDescription = Err.Description
    ' 解説: この限定範囲のCOMエラーを検査できるようにします。
    On Error GoTo 0
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Service query failed: " & errorNumber & " " & errorDescription
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐を終了します。
End If
' 解説: この限定範囲のCOMエラーを検査できるようにします。
On Error GoTo 0
' 解説: 取得結果またはエラーをコンソールへ表示します。
WScript.Echo "Service query succeeded; count=" & services.Count
