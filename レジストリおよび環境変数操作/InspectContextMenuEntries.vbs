' ファイル概要: シェルのコンテキストメニュー登録キーを列挙する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectContextMenuEntries.vbs
' 実行結果: シェルのコンテキストメニュー登録キーを列挙する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, names, subKey, item, child
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
For Each item In Array("*\shell", "Directory\shell", "Directory\Background\shell")
    result = registry.EnumKey(&H80000000, item, names)
    WScript.Echo "HKCR\" & item
    If result = 0 And IsArray(names) Then
        For Each child In names
            WScript.Echo "  " & child
        Next
    Else
        WScript.Echo "  No entries returned; WMI code " & result
    End If
Next
