' 必要なWMIクラスへのアクセス可否を照会し、互換性確認の簡単な結果を出します。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: service, services, relations
Dim service, services, relations
' 解説: この限定範囲のCOMエラーを検査できるようにします。
On Error Resume Next
' 解説: ローカルWMI名前空間に接続し、照会を行うオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Err.Number <> 0 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "WMI namespace unavailable: " & Err.Description
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐を終了します。
End If
' 解説: 直前に確認済みのエラー状態をクリアします。
Err.Clear
' 解説: サービス情報を読み取り専用で照会します。
Set services = service.ExecQuery("SELECT Name FROM Win32_Service")
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Err.Number <> 0 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Win32_Service query unavailable: " & Err.Description
    ' 解説: スクリプトを終了し、終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐を終了します。
End If
' 解説: 取得結果またはエラーをコンソールへ表示します。
WScript.Echo "Win32_Service available; count=" & services.Count
' 解説: 直前に確認済みのエラー状態をクリアします。
Err.Clear
' 解説: サービス情報を読み取り専用で照会します。
Set relations = service.ExecQuery("SELECT Antecedent, Dependent FROM Win32_DependentService")
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Err.Number <> 0 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Win32_DependentService unavailable: " & Err.Description
    ' 解説: 直前に確認済みのエラー状態をクリアします。
    Err.Clear
' 解説: 先行条件に該当しない場合の処理です。
Else
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Win32_DependentService available; count=" & relations.Count
' 解説: 条件分岐を終了します。
End If
' 解説: この限定範囲のCOMエラーを検査できるようにします。
On Error GoTo 0
