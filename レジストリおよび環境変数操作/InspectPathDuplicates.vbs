' ファイル概要: 現在のプロセスPATHの重複を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectPathDuplicates.vbs
' 実行結果: 現在のプロセスPATHの重複を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, pathValue, entries, seen, entry, normalized, foundDuplicate
Set shell = CreateObject("WScript.Shell")
pathValue = shell.Environment("PROCESS")("PATH")
Set seen = CreateObject("Scripting.Dictionary")
seen.CompareMode = vbTextCompare
foundDuplicate = False
entries = Split(pathValue, ";")
For Each entry In entries
    normalized = Trim(entry)
    If Len(normalized) > 0 Then
        If seen.Exists(normalized) Then
            WScript.Echo "Duplicate PATH entry: " & normalized
            foundDuplicate = True
        Else
            seen.Add normalized, True
        End If
    End If
Next
If Not foundDuplicate Then WScript.Echo "No duplicate entries found in the current process PATH."
