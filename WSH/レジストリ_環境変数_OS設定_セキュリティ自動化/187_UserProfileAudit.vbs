' ローカルプロファイルの合計・ロード済み件数だけを表示します。ユーザー名は出力しません。
Option Explicit
Dim service, profiles, profile, total, loaded
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set profiles = service.ExecQuery("SELECT Loaded, Special FROM Win32_UserProfile")
total = 0
loaded = 0
For Each profile In profiles
    If Not profile.Special Then
        total = total + 1
        If profile.Loaded Then loaded = loaded + 1
    End If
Next
WScript.Echo "NonSpecialProfiles=" & total & " Loaded=" & loaded
