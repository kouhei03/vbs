' 概要: ごみ箱の内容と件数を表示します。完全削除は確認語を指定した場合だけ実行します。
' 実行方法: Usage: cscript //nologo 029_EmptyRecycleBin.vbs [EMPTY-RECYCLE-BIN]
' 処理結果: ごみ箱の内容表示／確認語指定時の削除要求。処理結果や状態はコンソールに表示されます。
' コード解説: スクリプト引数を使って対象を特定し、検証後に処理します。
Option Explicit
Dim shell, recycleBin, items, item, confirm, count
confirm = ""
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count > 1 Then
    WScript.Echo "Usage: cscript //nologo 029_EmptyRecycleBin.vbs [EMPTY-RECYCLE-BIN]"
    WScript.Quit 2
End If
If WScript.Arguments.Count = 1 Then confirm = UCase(WScript.Arguments(0))
Set shell = CreateObject("Shell.Application")
Set recycleBin = shell.Namespace(10)
Set items = recycleBin.Items
count = items.Count
' 主処理: 対象の一覧を1件ずつ処理します。
For Each item In items
    WScript.Echo item.Name
Next
WScript.Echo "Items=" & count
If confirm = "EMPTY-RECYCLE-BIN" Then
    items.InvokeVerb "delete"
    WScript.Echo "Delete requested; Windows may still display a confirmation dialog."
Else
    WScript.Echo "Preview only; pass EMPTY-RECYCLE-BIN to request deletion."
End If
