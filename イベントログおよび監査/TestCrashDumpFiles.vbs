Option Explicit

Dim fileSystem, shell, windowsPath, dumpPath, miniDumpFolder, folder
Dim fileItem, found
Set fileSystem = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
windowsPath = shell.ExpandEnvironmentStrings("%SystemRoot%")

dumpPath = fileSystem.BuildPath(windowsPath, "MEMORY.DMP")
If fileSystem.FileExists(dumpPath) Then
    WScript.Echo "Memory dump exists: " & dumpPath
Else
    WScript.Echo "No MEMORY.DMP found at the standard path."
End If

miniDumpFolder = fileSystem.BuildPath(windowsPath, "Minidump")
found = False
If fileSystem.FolderExists(miniDumpFolder) Then
    Set folder = fileSystem.GetFolder(miniDumpFolder)
    For Each fileItem In folder.Files
        If LCase(fileSystem.GetExtensionName(fileItem.Name)) = "dmp" Then
            found = True
            WScript.Echo "Minidump exists: " & fileItem.Path
        End If
    Next
End If
If Not found Then WScript.Echo "No minidump files found in the standard folder."
