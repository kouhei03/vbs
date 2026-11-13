' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（network）。主な用途: ネットワーク情報オブジェクト。
Dim network
' 解説: WScript.Network（コンピューター名や割り当てドライブを参照するオブジェクト）を生成し、変数に格納します。
Set network = CreateObject("WScript.Network")
' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Computer name: " & network.ComputerName
WScript.Echo "Computer name: " & network.ComputerName
