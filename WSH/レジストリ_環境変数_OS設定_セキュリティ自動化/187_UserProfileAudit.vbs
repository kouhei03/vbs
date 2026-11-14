' 概要: ローカルプロファイルの合計・ロード済み件数だけを表示します。ユーザー名は出力しません。
' 実行方法: cscript //nologo 187_UserProfileAudit.vbs
' 処理結果: プロファイル件数とロード済み件数。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, profiles, profile, total, loaded
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
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
