' ファイル概要: ユーザープロファイルのSIDとハイブの状態を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectUserProfileHives.vbs
' 実行結果: ユーザープロファイルのSIDとハイブの状態を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, profiles, profile, fso, hivePath, hiveExists
Set service = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set fso = CreateObject("Scripting.FileSystemObject")
Set profiles = service.ExecQuery("SELECT SID, LocalPath, Loaded FROM Win32_UserProfile WHERE Special = FALSE")
For Each profile In profiles
    hivePath = profile.LocalPath & "\NTUSER.DAT"
    hiveExists = fso.FileExists(hivePath)
    WScript.Echo "SID=" & profile.SID & " Loaded=" & profile.Loaded & " NTUSER.DAT exists=" & hiveExists
Next
WScript.Echo "No user hive was loaded, edited, or removed."
