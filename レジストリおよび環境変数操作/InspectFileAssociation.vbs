Option Explicit
Dim registry, extension, progId, result, hive
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo InspectFileAssociation.vbs <.extension>"
    WScript.Quit 2
End If
extension = WScript.Arguments(0)
If Left(extension, 1) <> "." Or InStr(extension, "\") > 0 Then
    WScript.Echo "Provide a single extension beginning with a period."
    WScript.Quit 2
End If
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetStringValue(&H80000000, extension, "", progId)
If result = 0 Then WScript.Echo "HKCR default ProgID=" & progId
result = registry.GetStringValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & extension & "\UserChoice", "ProgId", progId)
If result = 0 Then
    WScript.Echo "UserChoice ProgID=" & progId
Else
    WScript.Echo "UserChoice ProgID is not available; WMI return code: " & result
End If
