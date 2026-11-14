' ファイル概要: レジストリ文字列の書き込み対象と文字数を確認する（書き込みなし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo PreviewRegistryStringWrite.vbs <hive> <subkey> <value-name> <value>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: レジストリ文字列の書き込み対象と文字数を確認する（書き込みなし）処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim args, hiveName, subKey, valueName, proposedValue
Set args = WScript.Arguments
If args.Count <> 4 Then
    WScript.Echo "Usage: cscript //nologo PreviewRegistryStringWrite.vbs <hive> <subkey> <value-name> <value>"
    WScript.Quit 2
End If
hiveName = UCase(args(0))
subKey = args(1)
valueName = args(2)
proposedValue = args(3)
If InStr("|HKCU|HKLM|HKCR|HKU|HKCC|", "|" & hiveName & "|") = 0 Then
    WScript.Echo "Unsupported registry hive: " & hiveName
    WScript.Quit 2
End If
WScript.Echo "READ-ONLY PREVIEW; no registry value was written."
WScript.Echo "Target: " & hiveName & "\" & subKey & "\" & valueName
WScript.Echo "Proposed REG_SZ character count: " & Len(proposedValue)
