' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: ユーザープロファイル情報とNTUSER.DATの存在確認に使う変数を宣言します。
Dim service, profiles, profile, fso, hivePath, hiveExists
Set service = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set fso = CreateObject("Scripting.FileSystemObject")
' 解説: 登録済みローカルプロファイルを列挙し、ハイブのパスとロード状態だけを確認します。
Set profiles = service.ExecQuery("SELECT SID, LocalPath, Loaded FROM Win32_UserProfile WHERE Special = FALSE")
For Each profile In profiles
    hivePath = profile.LocalPath & "\NTUSER.DAT"
    hiveExists = fso.FileExists(hivePath)
    WScript.Echo "SID=" & profile.SID & " Loaded=" & profile.Loaded & " NTUSER.DAT exists=" & hiveExists
Next
WScript.Echo "No user hive was loaded, edited, or removed."
