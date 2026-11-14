' ファイル概要: IFEOのDebugger値の有無を確認する（値は非表示）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectIFEOBlockSettings.vbs <image-name.exe>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: IFEOのDebugger値の有無を確認する（値は非表示）処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, imageName, subKey, result, names, types, i
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo InspectIFEOBlockSettings.vbs <image-name.exe>"
    WScript.Quit 2
End If
imageName = WScript.Arguments(0)
If InStr(imageName, "\") > 0 Or InStr(imageName, "/") > 0 Then
    WScript.Echo "Provide only an image file name."
    WScript.Quit 2
End If
subKey = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\" & imageName
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' [処理解説] 指定キーの値名と型を列挙します。値データはこの操作では取得しません。
result = registry.EnumValues(&H80000002, subKey, names, types)
If result <> 0 Then
    WScript.Echo "IFEO key unavailable; WMI return code: " & result
    WScript.Quit 1
End If
WScript.Echo "Target: " & imageName
If IsArray(names) Then
    For i = 0 To UBound(names)
        If StrComp(names(i), "Debugger", vbTextCompare) = 0 Then
            WScript.Echo "Debugger value is configured; data withheld."
            WScript.Quit 0
        End If
    Next
End If
WScript.Echo "No Debugger value found."
