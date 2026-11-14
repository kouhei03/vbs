Option Explicit

Dim service, users, profiles, shares, user, profile, share
Dim userCount, disabledCount, lockedCount, profileCount, shareCount
Dim fileSystem, outputPath, reportHtml, reportFile

If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo CreateSecurityAuditSummary.vbs <new-html-path>"
    WScript.Quit 2
End If
outputPath = WScript.Arguments(0)
Set fileSystem = CreateObject("Scripting.FileSystemObject")
If fileSystem.FileExists(outputPath) Then
    WScript.Echo "Output file already exists; refusing to overwrite: " & outputPath
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
Set users = service.ExecQuery( _
    "SELECT Disabled, Lockout FROM Win32_UserAccount WHERE LocalAccount=True")
Set profiles = service.ExecQuery( _
    "SELECT SID FROM Win32_UserProfile WHERE Special=False")
Set shares = service.ExecQuery("SELECT Name FROM Win32_Share")

userCount = 0
disabledCount = 0
lockedCount = 0
For Each user In users
    userCount = userCount + 1
    If user.Disabled Then disabledCount = disabledCount + 1
    If user.Lockout Then lockedCount = lockedCount + 1
Next
profileCount = 0
For Each profile In profiles
    profileCount = profileCount + 1
Next
shareCount = 0
For Each share In shares
    shareCount = shareCount + 1
Next

reportHtml = "<!doctype html>" & vbCrLf & _
    "<html><head><meta charset=""utf-8""><title>Local Security Summary</title></head><body>" & vbCrLf & _
    "<h1>Local Security Inventory Summary</h1>" & vbCrLf & _
    "<p>Generated: " & ServerTime() & "</p>" & vbCrLf & _
    "<ul><li>Local user accounts: " & userCount & "</li>" & _
    "<li>Disabled accounts: " & disabledCount & "</li>" & _
    "<li>Locked accounts: " & lockedCount & "</li>" & _
    "<li>Non-special user profiles: " & profileCount & "</li>" & _
    "<li>Local shares: " & shareCount & "</li></ul>" & vbCrLf & _
    "<p>Summary counts only; no names, paths, credentials, or ACL entries are included.</p>" & _
    "</body></html>"

On Error Resume Next
Set reportFile = fileSystem.CreateTextFile(outputPath, False, True)
If Err.Number <> 0 Then
    WScript.Echo "Could not create report: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0
reportFile.Write reportHtml
reportFile.Close
WScript.Echo "Summary report written to: " & outputPath

Function ServerTime()
    ServerTime = CStr(Now)
End Function
