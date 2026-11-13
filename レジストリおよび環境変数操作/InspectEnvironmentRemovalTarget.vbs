' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: WScript.Shell、ユーザー環境変数、および入力名を保持します。
Dim shell, environment, variableName, entry, separator, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo InspectEnvironmentRemovalTarget.vbs <variable-name>"
    WScript.Quit 2
End If
variableName = WScript.Arguments(0)
Set shell = CreateObject("WScript.Shell")
Set environment = shell.Environment("USER")
' 解説: ユーザー環境変数の削除対象が存在するかを確認し、削除や通知は行いません。
found = False
For Each entry In environment
    separator = InStr(entry, "=")
    If separator > 0 Then
        If StrComp(Left(entry, separator - 1), variableName, vbTextCompare) = 0 Then found = True
    End If
Next
If found Then
    WScript.Echo "Variable exists in the current user environment: " & variableName
Else
    WScript.Echo "Variable is not present in the current user environment: " & variableName
End If
WScript.Echo "No environment variable was removed."
