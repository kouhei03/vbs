' ファイル概要: ユーザープロファイルのパスと状態を一覧表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo ListUserProfiles.vbs
' 実行結果: ユーザープロファイルのパスと状態を一覧表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, profiles, profile
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set profiles = service.ExecQuery( _
    "SELECT SID, LocalPath, Loaded, Special FROM Win32_UserProfile " & _
    "WHERE Special=False")

WScript.Echo "SID" & vbTab & "Profile path" & vbTab & "Loaded"
For Each profile In profiles
    WScript.Echo profile.SID & vbTab & profile.LocalPath & vbTab & profile.Loaded
Next
