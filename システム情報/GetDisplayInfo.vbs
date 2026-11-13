' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、controllers、controller）。主な用途: WMIサービス接続、表示デバイス一覧、表示デバイス情報。
Dim service, controllers, controller
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのビデオコントローラー（Win32_VideoController）を読み取り専用で照会し、Name, CurrentHorizontalResolution, CurrentVerticalResolutionを取得します。
Set controllers = service.ExecQuery( _
    "SELECT Name, CurrentHorizontalResolution, CurrentVerticalResolution " & _
    "FROM Win32_VideoController")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each controller In controllers
For Each controller In controllers
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Display adapter: " & controller.Name
    WScript.Echo "Display adapter: " & controller.Name
    ' 解説: 条件「Not IsNull(controller.CurrentHorizontalResolution) And Not IsNull(controller.CurrentVerticalResolution) Then」を評価し、成立した場合に続く分岐処理を実行します。
    If Not IsNull(controller.CurrentHorizontalResolution) And _
        Not IsNull(controller.CurrentVerticalResolution) Then
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Reported resolution: " & controller.CurrentHorizontalResolution & " x " & controller.CurrentVerticalResolution
        WScript.Echo "Reported resolution: " & _
            controller.CurrentHorizontalResolution & " x " & _
            controller.CurrentVerticalResolution
    ' 解説: 直前までの条件がどれも成立しない場合の代替処理に入ります。
    Else
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Current resolution is not reported by this adapter."
        WScript.Echo "Current resolution is not reported by this adapter."
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
