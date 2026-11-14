' 固定回数のコンソール進捗表示をデモします。
Option Explicit
Dim i, count, bar
count = 20
For i = 1 To count
    bar = String(i, "#") & String(count - i, "-")
    WScript.StdOut.Write vbCr & "[" & bar & "] " & CInt(i * 100 / count) & "%"
    WScript.Sleep 100
Next
WScript.Echo
