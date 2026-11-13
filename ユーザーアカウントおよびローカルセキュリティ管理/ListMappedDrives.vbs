' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（network、mappings、index）。主な用途: ネットワーク情報オブジェクト、ドライブ割り当て一覧。
Dim network, mappings, index
' 解説: WScript.Network（コンピューター名や割り当てドライブを参照するオブジェクト）を生成し、変数に格納します。
Set network = CreateObject("WScript.Network")
' 解説: 現在のユーザーに割り当てられているネットワークドライブ一覧を読み取ります。割り当ては変更しません。
Set mappings = network.EnumNetworkDrives

' 解説: 条件「mappings.Count = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If mappings.Count = 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "No network drives are mapped."
    WScript.Echo "No network drives are mapped."
' 解説: 直前までの条件がどれも成立しない場合の代替処理に入ります。
Else
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Drive" & vbTab & "Network path"
    WScript.Echo "Drive" & vbTab & "Network path"
    ' 解説: 指定範囲の繰り返し処理を開始します: For index = 0 To mappings.Count - 1 Step 2
    For index = 0 To mappings.Count - 1 Step 2
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: mappings.Item(index) & vbTab & mappings.Item(index + 1)
        WScript.Echo mappings.Item(index) & vbTab & mappings.Item(index + 1)
    ' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
    Next
' 解説: 条件分岐ブロックを終了します。
End If
