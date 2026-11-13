' 解説: DirectXバージョンをStdRegProvで読み取ります。レジストリ値は変更しません。
Option Explicit
' 解説: レジストリプロバイダー、戻り値、および取得値を保持します。
Dim registry, result, value
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: 指定キーの値を読み取り専用で取得します。
result = registry.GetStringValue(&H80000002, "SOFTWARE\Microsoft\DirectX", "Version", value)
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: result = 0
If result = 0 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Version=" & value
' 解説: 先行条件が不成立の場合に追加条件を評価します: result = 2
ElseIf result = 2 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Version is not set."
' 解説: 先行条件がいずれも不成立の場合の代替処理へ進みます。
Else
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Registry read failed with WMI return code: " & result
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
