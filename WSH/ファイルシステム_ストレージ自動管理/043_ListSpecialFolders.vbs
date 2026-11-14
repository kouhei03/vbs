' WScript.Shellが提供するユーザー特殊フォルダーのパスを表示します。
Option Explicit
Dim shell, folders, name
Set shell = CreateObject("WScript.Shell")
Set folders = shell.SpecialFolders
For Each name In Array("AllUsersDesktop", "AllUsersStartMenu", "AllUsersPrograms", "AllUsersStartup", _
    "Desktop", "Favorites", "Fonts", "MyDocuments", "NetHood", "PrintHood", _
    "Programs", "Recent", "SendTo", "StartMenu", "Startup", "Templates")
    WScript.Echo name & vbTab & folders(name)
Next
