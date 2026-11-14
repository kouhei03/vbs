' normal/debugスイッチに応じてログ詳細度を切り替える例です。
Option Explicit
Dim mode, verbose
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 090_VerboseModeDemo.vbs <normal|debug>"
    WScript.Quit 2
End If
mode = LCase(WScript.Arguments(0))
If mode <> "normal" And mode <> "debug" Then WScript.Quit 2
verbose = (mode = "debug")
WScript.Echo "Mode=" & mode
If verbose Then
    WScript.Echo "Timestamp=" & Now
    WScript.Echo "Script=" & WScript.ScriptFullName
    WScript.Echo "ArgumentCount=" & WScript.Arguments.Count
End If
