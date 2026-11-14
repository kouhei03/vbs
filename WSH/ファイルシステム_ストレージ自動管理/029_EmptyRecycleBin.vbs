' ごみ箱の内容と件数を表示します。完全削除は確認語を指定した場合だけ実行します。
Option Explicit
Dim shell, recycleBin, items, item, confirm, count
confirm = ""
If WScript.Arguments.Count > 1 Then
    WScript.Echo "Usage: cscript //nologo 029_EmptyRecycleBin.vbs [EMPTY-RECYCLE-BIN]"
    WScript.Quit 2
End If
If WScript.Arguments.Count = 1 Then confirm = UCase(WScript.Arguments(0))
Set shell = CreateObject("Shell.Application")
Set recycleBin = shell.Namespace(10)
Set items = recycleBin.Items
count = items.Count
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
