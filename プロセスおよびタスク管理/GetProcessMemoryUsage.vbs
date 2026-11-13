' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、processes、process）。主な用途: WMIサービス接続、プロセス一覧、プロセス情報。
Dim service, processes, process
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの実行中プロセス（Win32_Process）を読み取り専用で照会し、Name, ProcessId, WorkingSetSize, VirtualSizeを取得します。
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, WorkingSetSize, VirtualSize FROM Win32_Process")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Process name" & vbTab & "PID" & vbTab & "Working set (bytes)" & vbTab & "Virtual size (bytes)"
WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "Working set (bytes)" & vbTab & "Virtual size (bytes)"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each process In processes
For Each process In processes
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: process.Name & vbTab & process.ProcessId & vbTab & ByteCount(process.WorkingSetSize) & vbTab & ByteCount(process.VirtualSize)
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        ByteCount(process.WorkingSetSize) & vbTab & _
        ByteCount(process.VirtualSize)
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next

' 解説: 再利用可能な関数を定義します。関数名と引数: ByteCount(value)
Function ByteCount(value)
    ' 解説: 条件「IsNull(value) Or IsEmpty(value) Then」を評価し、成立した場合に続く分岐処理を実行します。
    If IsNull(value) Or IsEmpty(value) Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: ByteCount = "N/A"
        ByteCount = "N/A"
    ' 解説: 直前までの条件がどれも成立しない場合の代替処理に入ります。
    Else
        ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: ByteCount = FormatNumber(CDbl(value), 0)
        ByteCount = FormatNumber(CDbl(value), 0)
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 関数定義を終了します。
End Function
