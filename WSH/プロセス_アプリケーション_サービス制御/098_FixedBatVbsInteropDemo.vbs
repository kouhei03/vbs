' 一時BATを生成して固定のver照会だけを同期実行し、終了後に削除する連携デモです。
Option Explicit
Dim fso, shell, batPath, bat, exitCode
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
batPath = fso.BuildPath(fso.GetSpecialFolder(2).Path, "wsh-098-fixed-demo.bat")
If fso.FileExists(batPath) Then
    WScript.Echo "Demo file already exists; inspect before removal: " & batPath
    WScript.Quit 1
End If
Set bat = fso.CreateTextFile(batPath, False, False)
bat.WriteLine "@echo off"
bat.WriteLine "ver"
bat.Close
exitCode = shell.Run("cmd.exe /c """ & batPath & """", 1, True)
If fso.FileExists(batPath) Then fso.DeleteFile batPath, True
WScript.Echo "Fixed BAT exit code=" & exitCode
If exitCode <> 0 Then WScript.Quit exitCode
