' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、processes、process、userName、domainName、result）。主な用途: WMIサービス接続、プロセス一覧、プロセス情報、対象ユーザー名、処理結果コード。
Dim service, processes, process, userName, domainName, result
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの実行中プロセス（Win32_Process）を読み取り専用で照会し、Name, ProcessIdを取得します。
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId FROM Win32_Process")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Process name" & vbTab & "PID" & vbTab & "Owner"
WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Owner"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each process In processes
For Each process In processes
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: userName = ""
    userName = ""
    ' 解説: 条件を確認しながら繰り返す監視/待機ループを開始します。
    domainName = ""
    ' 解説: プロセス所有者をWMIメソッドで照会し、戻り値コードと所有者情報を受け取ります。
    result = process.GetOwner(userName, domainName)
    ' 解説: 条件「result = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
    If result = 0 Then
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: process.Name & vbTab & process.ProcessId & vbTab & domainName & "\" & userName
        WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
            domainName & "\" & userName
    ' 解説: 直前までの条件がどれも成立しない場合の代替処理に入ります。
    Else
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: process.Name & vbTab & process.ProcessId & vbTab & "Unavailable (code " & result & ")"
        WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
            "Unavailable (code " & result & ")"
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
