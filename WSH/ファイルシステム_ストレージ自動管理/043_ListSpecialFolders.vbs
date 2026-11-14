' 概要: WScript.Shellが提供するユーザー特殊フォルダーのパスを表示します。
' 実行方法: cscript //nologo 043_ListSpecialFolders.vbs
' 処理結果: WSH特殊フォルダーのパス表示。処理結果や状態はコンソールに表示されます。
' コード解説: WScript.Shellを使ってWindowsの設定や状態を照会します。
Option Explicit
Dim shell, folders, name
Set shell = CreateObject("WScript.Shell")
Set folders = shell.SpecialFolders
' 主処理: 対象の一覧を1件ずつ処理します。
For Each name In Array("AllUsersDesktop", "AllUsersStartMenu", "AllUsersPrograms", "AllUsersStartup", _
    "Desktop", "Favorites", "Fonts", "MyDocuments", "NetHood", "PrintHood", _
    "Programs", "Recent", "SendTo", "StartMenu", "Startup", "Templates")
    WScript.Echo name & vbTab & folders(name)
Next
