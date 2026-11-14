' ファイル概要: ディスプレイアダプターと解像度を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetDisplayInfo.vbs
' 実行結果: ディスプレイアダプターと解像度を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, controllers, controller
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_VideoController を照会し、Name, CurrentHorizontalResolution, CurrentVerticalResolution を取得します。
Set controllers = service.ExecQuery( _
    "SELECT Name, CurrentHorizontalResolution, CurrentVerticalResolution " & _
    "FROM Win32_VideoController")

For Each controller In controllers
    WScript.Echo "Display adapter: " & controller.Name
    If Not IsNull(controller.CurrentHorizontalResolution) And _
        Not IsNull(controller.CurrentVerticalResolution) Then
        WScript.Echo "Reported resolution: " & _
            controller.CurrentHorizontalResolution & " x " & _
            controller.CurrentVerticalResolution
    Else
        WScript.Echo "Current resolution is not reported by this adapter."
    End If
Next
