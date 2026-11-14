' ファイル概要: ローカルディスクの空き容量をしきい値と比較する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 161_DiskSpaceWarningPreview.vbs <free-percent-threshold 1-50>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: ローカルディスクの空き容量をしきい値と比較する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim threshold, service, disks, disk, freePercent
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 161_DiskSpaceWarningPreview.vbs <free-percent-threshold 1-50>"
    WScript.Quit 2
End If
If Not ParsePercent(WScript.Arguments(0), threshold) Then
    WScript.Echo "Threshold must be an integer from 1 to 50."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set disks = service.ExecQuery( _
    "SELECT DeviceID, Size, FreeSpace FROM Win32_LogicalDisk WHERE DriveType=3")
For Each disk In disks
    If Not IsNull(disk.Size) Then
        If Not IsNull(disk.FreeSpace) Then
            If CDbl(disk.Size) > 0 Then
                freePercent = (CDbl(disk.FreeSpace) / CDbl(disk.Size)) * 100
                If freePercent < threshold Then
                    WScript.Echo "WARNING " & disk.DeviceID & " free=" & _
                        FormatNumber(freePercent, 1) & "%"
                Else
                    WScript.Echo "OK " & disk.DeviceID & " free=" & _
                        FormatNumber(freePercent, 1) & "%"
                End If
            End If
        End If
    End If
Next
WScript.Echo "Console preview only; no warning event was generated."

Function ParsePercent(text, value)
    Dim i, character
    ParsePercent = False
    If Len(text) = 0 Or Len(text) > 2 Then Exit Function
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
    Next
    value = CLng(text)
    ParsePercent = (value >= 1 And value <= 50)
End Function
