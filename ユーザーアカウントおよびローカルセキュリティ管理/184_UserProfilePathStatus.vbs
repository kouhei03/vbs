' ファイル概要: ユーザープロファイルのパスと状態を列挙する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 184_UserProfilePathStatus.vbs
' 実行結果: ユーザープロファイルのパスと状態を列挙する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, profiles, profile
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_UserProfile を照会し、SID, LocalPath, Loaded, Special を取得します。
Set profiles = service.ExecQuery( _
    "SELECT SID, LocalPath, Loaded, Special FROM Win32_UserProfile")
WScript.Echo "SID" & vbTab & "Path" & vbTab & "Loaded" & vbTab & "Special"
For Each profile In profiles
    WScript.Echo profile.SID & vbTab & profile.LocalPath & vbTab & _
        profile.Loaded & vbTab & profile.Special
Next
