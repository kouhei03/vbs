' 概要: Applicationログへ固定の情報イベントを1件記録します。Security監査ログには書き込みません。
' 実行方法: cscript //nologo 177_CustomApplicationEvent.vbs
' 処理結果: Applicationログへ固定の情報イベントを1件記録。処理結果や状態はコンソールに表示されます。
' コード解説: WScript.Shellを使ってWindowsの設定や状態を照会します。
Option Explicit
Dim shell
' 主処理: Applicationログへ固定の情報イベントを1件記録。
Set shell = CreateObject("WScript.Shell")
If Not shell.LogEvent(4, "WSH safe example: informational configuration audit marker.") Then
    WScript.Echo "Could not write an Application information event."
    WScript.Quit 1
End If
WScript.Echo "One informational event was written to the Application log."
