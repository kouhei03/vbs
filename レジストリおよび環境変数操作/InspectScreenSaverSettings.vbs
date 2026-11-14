' ファイル概要: スクリーンセーバーの有効状態・タイムアウト・保護設定を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectScreenSaverSettings.vbs
' 実行結果: スクリーンセーバーの有効状態・タイムアウト・保護設定を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, value, valueName
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
For Each valueName In Array("ScreenSaveActive", "ScreenSaveTimeOut", "ScreenSaverIsSecure", "SCRNSAVE.EXE")
    ' [処理解説] 指定した型のレジストリ値をStdRegProvから取得します。
    result = registry.GetStringValue(&H80000001, "Control Panel\Desktop", valueName, value)
    If result = 0 Then
        WScript.Echo valueName & "=" & value
    ElseIf result = 2 Then
        WScript.Echo valueName & "=(not set)"
    Else
        WScript.Echo valueName & " read failed; WMI return code: " & result
    End If
Next
