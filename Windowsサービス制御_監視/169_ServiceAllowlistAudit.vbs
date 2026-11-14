' 外部許可リストとローカルサービス名を比較し、不一致を表示します。削除や無効化はしません。
Option Explicit
Dim fso, inputFile, allowed, service, services, item, path, lineText, found, allowedName
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 169_ServiceAllowlistAudit.vbs <allowlist-file>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(path) Then
    WScript.Echo "Allowlist file not found: " & path
    WScript.Quit 2
End If
Set allowed = CreateObject("Scripting.Dictionary")
allowed.CompareMode = 1
Set inputFile = fso.OpenTextFile(path, 1, False)
Do Until inputFile.AtEndOfStream
    lineText = Trim(inputFile.ReadLine)
    If Len(lineText) > 0 Then allowed(lineText) = True
Loop
inputFile.Close
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery("SELECT Name, DisplayName, State FROM Win32_Service")
For Each item In services
    If Not allowed.Exists(item.Name) Then WScript.Echo "Not in allowlist: " & item.Name & vbTab & item.DisplayName & vbTab & item.State
Next
WScript.Echo "Audit completed; no service settings were changed."
