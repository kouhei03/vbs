' 概要: ローカルユーザープロファイルのLoaded状態とNTUSER.DATの存在だけを確認します。ハイブ編集なし。
' 実行方法: cscript //nologo 172_UserHiveInventory.vbs
' 処理結果: プロファイルとNTUSER.DAT存在確認（編集なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim fso, service, profiles, profile, hivePath
Set fso = CreateObject("Scripting.FileSystemObject")
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set profiles = service.ExecQuery("SELECT LocalPath, Loaded, Special FROM Win32_UserProfile")
For Each profile In profiles
    If Not profile.Special And Not IsNull(profile.LocalPath) Then
        hivePath = fso.BuildPath(profile.LocalPath, "NTUSER.DAT")
        WScript.Echo "Profile=" & profile.LocalPath & " Loaded=" & profile.Loaded & _
            " NTUSER.DAT=" & fso.FileExists(hivePath)
    End If
Next
