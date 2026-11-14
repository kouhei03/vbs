' ローカルユーザープロファイルのLoaded状態とNTUSER.DATの存在だけを確認します。ハイブ編集なし。
Option Explicit
Dim fso, service, profiles, profile, hivePath
Set fso = CreateObject("Scripting.FileSystemObject")
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set profiles = service.ExecQuery("SELECT LocalPath, Loaded, Special FROM Win32_UserProfile")
For Each profile In profiles
    If Not profile.Special And Not IsNull(profile.LocalPath) Then
        hivePath = fso.BuildPath(profile.LocalPath, "NTUSER.DAT")
        WScript.Echo "Profile=" & profile.LocalPath & " Loaded=" & profile.Loaded & _
            " NTUSER.DAT=" & fso.FileExists(hivePath)
    End If
Next
