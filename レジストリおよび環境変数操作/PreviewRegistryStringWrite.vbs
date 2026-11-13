' 解説: 明示的な変数宣言を必須にし、未宣言変数の誤使用を検出します。
Option Explicit
' 解説: 入力引数、レジストリハイブ名、キー、値名、および値の長さを保持します。
Dim args, hiveName, subKey, valueName, proposedValue
' 解説: 4引数で書き込み対象を受け取りますが、レジストリは変更しません。
Set args = WScript.Arguments
If args.Count <> 4 Then
    WScript.Echo "Usage: cscript //nologo PreviewRegistryStringWrite.vbs <hive> <subkey> <value-name> <value>"
    WScript.Quit 2
End If
' 解説: 対象を表示するためにハイブ名を大文字化し、残りの入力を取得します。
hiveName = UCase(args(0))
subKey = args(1)
valueName = args(2)
proposedValue = args(3)
' 解説: 読み取り専用ツールで扱う許可済みハイブだけを受け付けます。
If InStr("|HKCU|HKLM|HKCR|HKU|HKCC|", "|" & hiveName & "|") = 0 Then
    WScript.Echo "Unsupported registry hive: " & hiveName
    WScript.Quit 2
End If
' 解説: 書き込み内容は表示せず、対象と文字数だけをプレビューします。
WScript.Echo "READ-ONLY PREVIEW; no registry value was written."
WScript.Echo "Target: " & hiveName & "\" & subKey & "\" & valueName
WScript.Echo "Proposed REG_SZ character count: " & Len(proposedValue)
