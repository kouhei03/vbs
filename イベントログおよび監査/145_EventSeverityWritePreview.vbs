' ファイル概要: 情報・警告イベントの入力を確認する（書き込みなし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 145_EventSeverityWritePreview.vbs <Information|Warning> <message>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 情報・警告イベントの入力を確認する（書き込みなし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim severity, messageText
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 145_EventSeverityWritePreview.vbs <Information|Warning> <message>"
    WScript.Quit 2
End If
severity = WScript.Arguments(0)
If StrComp(severity, "Information", vbTextCompare) <> 0 And _
   StrComp(severity, "Warning", vbTextCompare) <> 0 Then
    WScript.Echo "Severity must be Information or Warning."
    WScript.Quit 2
End If
messageText = WScript.Arguments(1)
If Len(messageText) > 2048 Then
    WScript.Echo "Message is too long (maximum 2048 characters)."
    WScript.Quit 2
End If

WScript.Echo "READ-ONLY PREVIEW; no event was written."
WScript.Echo "Severity: " & severity
WScript.Echo "Message length: " & Len(messageText)
