Option Explicit
Dim fso, shell, tempPath, folder, file, ageDays, cutoff, found
If WScript.Arguments.Count > 1 Then WScript.Echo "Usage: cscript //nologo 089_TempCleanupPreview.vbs [age-days]": WScript.Quit 2
ageDays = 30
If WScript.Arguments.Count = 1 Then
    If Not IsNumeric(WScript.Arguments(0)) Then WScript.Echo "Age must be a number.": WScript.Quit 2
    ageDays = CLng(WScript.Arguments(0))
End If
If ageDays < 1 Or ageDays > 3650 Then WScript.Echo "Age must be 1..3650 days.": WScript.Quit 2
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
tempPath = shell.ExpandEnvironmentStrings("%TEMP%")
Set folder = fso.GetFolder(tempPath)
cutoff = DateAdd("d", -ageDays, Now)
found = False
For Each file In folder.Files
    If file.DateLastModified < cutoff Then
        WScript.Echo file.Path & " modified=" & file.DateLastModified & " bytes=" & file.Size
        found = True
    End If
Next
If Not found Then WScript.Echo "No old files found in TEMP root. No files were deleted."
