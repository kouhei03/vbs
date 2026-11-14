' 指定サービスのDelayedAutoStartレジストリ値を読み取り専用で表示します。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: このスクリプトで使用する変数を宣言します: serviceName, validator, shell, registryPath, delayedValue
Dim serviceName, validator, shell, registryPath, delayedValue
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "Usage: cscript //nologo 109_DelayedAutoStartStatus.vbs <service-name>"
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
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set shell = CreateObject("WScript.Shell")
' 解説: 入力値、取得値、または計算結果を変数に格納します。
registryPath = "HKLM\SYSTEM\CurrentControlSet\Services" & serviceName & "\DelayedAutoStart"
' 解説: この限定範囲のCOMエラーを検査できるようにします。
On Error Resume Next
' 解説: 入力値、取得値、または計算結果を変数に格納します。
delayedValue = shell.RegRead(registryPath)
' 解説: 入力または取得結果を検証し、必要な分岐を選択します。
If Err.Number <> 0 Then
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "DelayedAutoStart value is absent or unreadable; this alone does not establish startup mode."
    ' 解説: 直前に確認済みのエラー状態をクリアします。
    Err.Clear
' 解説: 先行条件に該当しない場合の処理です。
Else
    ' 解説: 取得結果またはエラーをコンソールへ表示します。
    WScript.Echo "DelayedAutoStart=" & delayedValue & " (1 usually indicates delayed automatic start)"
' 解説: 条件分岐を終了します。
End If
' 解説: この限定範囲のCOMエラーを検査できるようにします。
On Error GoTo 0
