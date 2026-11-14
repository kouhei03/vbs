' 終了コードを定義済みメッセージへ対応づけます（外部コマンドは実行しません）。
Option Explicit
Dim code
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 092_ExitCodeMessageMap.vbs <exit-code>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(0)) Then WScript.Quit 2
If CDbl(WScript.Arguments(0)) < -32768 Or CDbl(WScript.Arguments(0)) > 32767 Then WScript.Quit 2
code = CLng(WScript.Arguments(0))
Select Case code
    Case 0: WScript.Echo "Success: operation completed."
    Case 1: WScript.Echo "General failure: inspect the operation log."
    Case 2: WScript.Echo "Invalid arguments or configuration."
    Case 3: WScript.Echo "Requested item was not found."
    Case Else: WScript.Echo "Unmapped exit code: " & code
End Select
