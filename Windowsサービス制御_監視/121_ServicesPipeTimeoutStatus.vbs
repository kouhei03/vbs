' ServicesPipeTimeoutの現在値を読み取ります。レジストリ変更はしません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: shell, value, registryPath
Dim shell, value, registryPath
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 入力値、取得値、または計算結果を変数に格納します。
registryPath = "HKLM\SYSTEM\CurrentControlSet\Control\ServicesPipeTimeout"
' 解説: この限定範囲のCOMエラーを検査できるようにします。
On Error Resume Next
' 解説: 入力値、取得値、または計算結果を変数に格納します。
value = shell.RegRead(registryPath)
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Err.Number <> 0 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "ServicesPipeTimeout is not explicitly set or cannot be read."
    ' 解説: 直前に確認済みのエラー状態をクリアします。
    Err.Clear
' 解説: 先行条件に該当しない場合の処理です。
Else
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "ServicesPipeTimeout (milliseconds)=" & value
' 解説: 条件分岐を終了します。
End If
' 解説: この限定範囲のCOMエラーを検査できるようにします。
On Error GoTo 0
