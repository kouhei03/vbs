' Edge/Chromeポリシーキーの主要な値が設定されているかだけ確認します。
Option Explicit
Dim shell, roots, root, names, name, value, found
Set shell = CreateObject("WScript.Shell")
roots = Array( _
    "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge\", _
    "HKEY_CURRENT_USER\SOFTWARE\Policies\Microsoft\Edge\", _
    "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Google\Chrome\", _
    "HKEY_CURRENT_USER\SOFTWARE\Policies\Google\Chrome\")
names = Array("HomepageLocation", "RestoreOnStartup", "DefaultSearchProviderEnabled")
For Each root In roots
    For Each name In names
        On Error Resume Next
        value = shell.RegRead(root & name)
        If Err.Number = 0 Then
            WScript.Echo root & name & "=(configured)"
        End If
        Err.Clear
        On Error GoTo 0
    Next
Next
WScript.Echo "Policy values were not changed; content values are omitted."
