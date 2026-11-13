' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: 現在のPATH、重複判定用辞書、および分割した要素を保持します。
Dim shell, pathValue, entries, seen, entry, normalized, foundDuplicate
Set shell = CreateObject("WScript.Shell")
pathValue = shell.Environment("PROCESS")("PATH")
Set seen = CreateObject("Scripting.Dictionary")
seen.CompareMode = vbTextCompare
foundDuplicate = False
' 解説: PATHをセミコロンで分け、空白を除いて大文字小文字を無視した重複を検出します。
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
