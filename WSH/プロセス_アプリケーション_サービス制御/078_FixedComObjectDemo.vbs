' 固定のFileSystemObject COMクラスを生成し、TEMPパス照会メソッドだけを呼び出します。
Option Explicit
Dim fso, tempFolder
Set fso = CreateObject("Scripting.FileSystemObject")
Set tempFolder = fso.GetSpecialFolder(2)
WScript.Echo "COM object=" & TypeName(fso)
WScript.Echo "TEMP=" & tempFolder.Path
