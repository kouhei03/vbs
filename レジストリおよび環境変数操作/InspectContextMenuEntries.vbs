' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: コンテキストメニュー登録先のサブキーを列挙するための変数を宣言します。
Dim registry, result, names, subKey, item, child
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: ファイルとディレクトリの代表的なシェル拡張登録先を読み取ります。
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
