Option Explicit
Dim percent
For percent=0 To 100 Step 10
    WScript.StdOut.Write "Progress: " & percent & "%" & vbCr
    WScript.Sleep 100
Next
WScript.StdOut.WriteLine "Progress: 100%"
