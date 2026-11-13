' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、processes、process）。主な用途: WMIサービス接続、プロセス一覧、プロセス情報。
Dim service, processes, process
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの実行中プロセス（Win32_Process）を読み取り専用で照会し、Name, ProcessId, UserModeTime, KernelModeTimeを取得します。
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, UserModeTime, KernelModeTime FROM Win32_Process")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Process name" & vbTab & "PID" & vbTab & "User CPU time (100ns)" & vbTab & "Kernel CPU time (100ns)"
WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "User CPU time (100ns)" & vbTab & "Kernel CPU time (100ns)"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each process In processes
For Each process In processes
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: process.Name & vbTab & process.ProcessId & vbTab & process.UserModeTime & vbTab & process.KernelModeTime
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        process.UserModeTime & vbTab & process.KernelModeTime
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
