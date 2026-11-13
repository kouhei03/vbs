' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: IFEO対象のイメージ名、Debugger値の有無、および照会結果を保持します。
Dim registry, imageName, subKey, result, names, types, i
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
' 解説: IFEO配下にDebugger値があるかだけを確認し、値データは表示しません。
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
