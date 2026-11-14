' ファイル概要: IEとEdgeのホームページ設定を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectBrowserHomePages.vbs
' 実行結果: IEとEdgeのホームページ設定を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, value, item, subKey, valueName
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
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
