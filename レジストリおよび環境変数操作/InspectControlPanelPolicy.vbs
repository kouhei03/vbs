' ファイル概要: 現在のユーザーのコントロールパネル制限値を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectControlPanelPolicy.vbs
' 実行結果: 現在のユーザーのコントロールパネル制限値を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, restricted
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetDWORDValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel", restricted)
If result = 0 Then
    WScript.Echo "NoControlPanel=" & restricted
ElseIf result = 2 Then
    WScript.Echo "NoControlPanel is not set."
Else
    WScript.Echo "Registry read failed with WMI return code: " & result
    WScript.Quit 1
End If
