' ファイル概要: 隠しファイルと拡張子の表示設定を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectExplorerDisplaySettings.vbs
' 実行結果: 隠しファイルと拡張子の表示設定を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, settingName, settingValue
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
For Each settingName In Array("Hidden", "HideFileExt", "ShowSuperHidden")
    ' [処理解説] 指定した型のレジストリ値をStdRegProvから取得します。
    result = registry.GetDWORDValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", settingName, settingValue)
    If result = 0 Then
        WScript.Echo settingName & "=" & settingValue
    ElseIf result = 2 Then
        WScript.Echo settingName & "=(not set)"
    Else
        WScript.Echo settingName & " read failed; WMI return code: " & result
    End If
Next
