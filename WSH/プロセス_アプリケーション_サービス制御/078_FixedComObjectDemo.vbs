' 概要: 固定のFileSystemObject COMクラスを生成し、TEMPパス照会メソッドだけを呼び出します。
' 実行方法: cscript //nologo 078_FixedComObjectDemo.vbs
' 処理結果: FileSystemObjectの固定COM呼び出し例。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, tempFolder
' 主処理: FileSystemObjectの固定COM呼び出し例。
Set fso = CreateObject("Scripting.FileSystemObject")
Set tempFolder = fso.GetSpecialFolder(2)
WScript.Echo "COM object=" & TypeName(fso)
WScript.Echo "TEMP=" & tempFolder.Path
