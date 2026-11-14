' ファイル概要: ペアレンタルコントロール設定有無。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 134_ParentalControlStatus.vbs
' 実行結果: ペアレンタルコントロール設定有無処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, result, names, keyName
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
For Each keyName In Array("SOFTWARE\Microsoft\Parental Controls", "SOFTWARE\Policies\Microsoft\Windows\ParentalControls")
    ' [処理解説] 指定キーの直下のサブキーを列挙します。再帰的な走査ではありません。
    result = registry.EnumKey(&H80000002, keyName, names)
    If result = 0 Then
        WScript.Echo keyName & ": key present"
    ElseIf result = 2 Then
        WScript.Echo keyName & ": not present"
    Else
        WScript.Echo keyName & ": WMI code=" & result
    End If
Next
