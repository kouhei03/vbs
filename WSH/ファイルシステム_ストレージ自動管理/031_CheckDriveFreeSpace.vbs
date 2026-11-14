' 指定ドライブの空き容量を確認し、閾値を下回ると警告します。
Option Explicit
Dim fso, driveName, threshold, drive
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 031_CheckDriveFreeSpace.vbs <drive-letter> <minimum-free-MB>"
    WScript.Quit 2
End If
If Not IsNumeric(WScript.Arguments(1)) Then
    WScript.Echo "Minimum free-space threshold must be numeric."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
driveName = UCase(Replace(WScript.Arguments(0), ":", "")) & ":"
If Len(driveName) <> 2 Or CDbl(WScript.Arguments(1)) < 0 Then WScript.Quit 2
threshold = CDbl(WScript.Arguments(1)) * 1048576
If Not fso.DriveExists(driveName) Then
    WScript.Echo "Drive not found: " & driveName
    WScript.Quit 1
End If
Set drive = fso.GetDrive(driveName)
WScript.Echo "Drive=" & driveName & " FreeBytes=" & drive.FreeSpace
If drive.FreeSpace < threshold Then
    WScript.Echo "WARNING: free space is below " & WScript.Arguments(1) & " MB."
    WScript.Quit 1
End If
