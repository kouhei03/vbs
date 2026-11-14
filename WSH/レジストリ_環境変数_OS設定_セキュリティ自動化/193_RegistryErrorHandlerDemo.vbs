' 固定のデモ値を読み、存在しない場合とアクセスエラーを明示して処理します。
Option Explicit
Dim shell, path, value, errorNumber, errorText
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\WSH\SafeExamples\ErrorDemo"
On Error Resume Next
value = shell.RegRead(path)
errorNumber = Err.Number
errorText = Err.Description
Err.Clear
On Error GoTo 0
If errorNumber = 0 Then
    WScript.Echo "Value=" & value
ElseIf errorNumber = 1 Or errorNumber = -2147024894 Then
    WScript.Echo "Demo value is absent."
Else
    WScript.Echo "Registry read failed. Error=" & errorNumber & " " & errorText
    WScript.Quit 1
End If
