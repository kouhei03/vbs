' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: ブラウザー設定の候補キーと値を順に確認します。
Dim registry, result, value, item, subKey, valueName
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: IEのユーザー設定とEdgeの管理ポリシーのホームページ値を読み取ります。
For Each item In Array("IE", "Edge")
    If item = "IE" Then
        subKey = "Software\Microsoft\Internet Explorer\Main"
        valueName = "Start Page"
        result = registry.GetStringValue(&H80000001, subKey, valueName, value)
    Else
        subKey = "SOFTWARE\Policies\Microsoft\Edge"
        valueName = "HomepageLocation"
        result = registry.GetStringValue(&H80000002, subKey, valueName, value)
    End If
    If result = 0 Then
        WScript.Echo item & " " & valueName & "=" & value
    Else
        WScript.Echo item & " homepage value unavailable; WMI return code: " & result
    End If
Next
