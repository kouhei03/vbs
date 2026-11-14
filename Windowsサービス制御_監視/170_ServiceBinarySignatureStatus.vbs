' 指定サービスの実行ファイル署名状態をローカルで検証します。サービス設定は変更しません。
' 解説: 未宣言変数の誤使用を防ぐため、明示的な変数宣言を必須にします。
Option Explicit
' 解説: 入力、WMI照会、実行ファイルパス解析、PowerShell出力処理の変数を宣言します。
Dim serviceName, validator, service, services, item, found, rawPath, executablePath, exePosition, quotePosition
' 解説: PowerShell呼び出しと標準出力の読み取りに使う変数を宣言します。
Dim shell, psCommand, commandLine, process, line
' 解説: コマンドライン引数が1個かどうか確認します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 正しい実行形式を表示します。
    WScript.Echo "Usage: cscript //nologo 170_ServiceBinarySignatureStatus.vbs <service-name>"
    ' 解説: 引数エラーを示して終了します。
    WScript.Quit 2
' 解説: 引数チェックを終了します。
End If
' 解説: 対象サービス名を取得します。
serviceName = Trim(WScript.Arguments(0))
' 解説: サービス名の許可文字を検査する正規表現を作成します。
Set validator = New RegExp
' 解説: 任意コマンド構文がサービス名に混入しないよう文字を制限します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
' 解説: 不正なサービス名を拒否します。
If Not validator.Test(serviceName) Then
    ' 解説: 入力要件を表示します。
    WScript.Echo "Invalid service name."
    ' 解説: 不正な入力として終了します。
    WScript.Quit 2
' 解説: サービス名チェックを終了します。
End If
' 解説: ローカルWMI名前空間へ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: サービス名と実行パスを読み取り専用で取得します。
Set services = service.ExecQuery("SELECT Name, PathName FROM Win32_Service")
' 解説: 対象の初期状態を「未検出」にします。
found = False
' 解説: ローカルサービス一覧から指定サービスを探します。
For Each item In services
    ' 解説: サービス名を大文字小文字を区別せず照合します。
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        ' 解説: サービスが見つかったことを記録します。
        found = True
        ' 解説: 実行イメージのコマンドラインがNullでないことを確認してから取得します。
        If IsNull(item.PathName) Then
            ' 解説: Nullの場合はパスなしとして扱います。
            rawPath = ""
        ' 解説: 通常の文字列パスを取得します。
        Else
            ' 解説: 前後の空白を取り除いた実行イメージ文字列を保存します。
            rawPath = Trim(CStr(item.PathName))
        ' 解説: パス値の分岐を終了します。
        End If
        ' 解説: 抽出結果の初期値を空にします。
        executablePath = ""
        ' 解説: 引用符で始まる実行パスの形式を処理します。
        If Len(rawPath) > 0 And Left(rawPath, 1) = Chr(34) Then
            ' 解説: 最初に閉じる引用符の位置を探します。
            quotePosition = InStr(2, rawPath, Chr(34))
            ' 解説: 引用符内に実行パスがある場合、その部分だけを取り出します。
            If quotePosition > 2 Then executablePath = Mid(rawPath, 2, quotePosition - 2)
        ' 解説: 引用符で囲まれていない実行パス形式を処理します。
        Else
            ' 解説: 実行ファイル拡張子の終端を探します。
            exePosition = InStr(1, LCase(rawPath), ".exe", vbBinaryCompare)
            ' 解説: 拡張子を含むパス部分を取り出します。
            If exePosition > 0 Then executablePath = Left(rawPath, exePosition + 3)
        ' 解説: パス形式の分岐を終了します。
        End If
        ' 解説: 実行ファイルパスが取得できない場合は署名照会を行わず案内します。
        If Len(executablePath) = 0 Then
            ' 解説: 実行パスが一般的な.exe形式でないことを表示します。
            WScript.Echo "Could not extract an executable .exe path from: " & rawPath
            ' 解説: 対象サービスの処理を終了します。
            Exit For
        ' 解説: パスを取得できた場合は署名状態を照会します。
        Else
            ' 解説: PowerShell文字列リテラル用にシングルクォートをエスケープし、署名の状態と発行者だけを問い合わせます。
            ' 解説: ローカル環境変数を展開し、実際の実行ファイルパスに変換します。
            Set shell = CreateObject("WScript.Shell")
            ' 解説: 実行パス内の環境変数を現在のローカル環境値へ展開します。
            executablePath = shell.ExpandEnvironmentStrings(executablePath)
            ' 解説: PowerShell文字列リテラル用の引用符をエスケープして署名照会式を組み立てます。
            psCommand = "Get-AuthenticodeSignature -LiteralPath '" & Replace(executablePath, "'", "''") & "' | Select-Object Status,@{Name='Publisher';Expression={$_.SignerCertificate.Subject}} | Format-List"
            ' 解説: 引数を固定したローカル署名確認コマンドを組み立てます。
            commandLine = "powershell.exe -NoProfile -NonInteractive -Command " & Chr(34) & psCommand & Chr(34)
            ' 解説: PowerShellの署名状態確認を実行します。
            ' 解説: 実行結果を標準出力と標準エラーから取得します。
            Set process = shell.Exec(commandLine)
            ' 解説: 署名状態と発行者の出力を表示します。
            Do While Not process.StdOut.AtEndOfStream
                ' 解説: 標準出力から1行読み取ります。
                line = process.StdOut.ReadLine
                ' 解説: 読み取った結果を表示します。
                WScript.Echo line
            ' 解説: 標準出力の読み取りを続けます。
            Loop
            ' 解説: PowerShellのエラー出力を表示します。
            Do While Not process.StdErr.AtEndOfStream
                ' 解説: 標準エラーから1行読み取ります。
                line = process.StdErr.ReadLine
                ' 解説: エラー内容を表示します。
                WScript.Echo "ERROR: " & line
            ' 解説: 標準エラーの読み取りを続けます。
            Loop
            ' 解説: 署名照会が失敗した場合、終了コードを呼び出し元へ返します。
            If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
        ' 解説: パスの有無による分岐を終了します。
        End If
        ' 解説: 対象サービスが確定したため一覧走査を終了します。
        Exit For
    ' 解説: サービス名照合の分岐を終了します。
    End If
' 解説: 次のサービスを確認します。
Next
' 解説: 対象サービスが見つからなかった場合はエラーを返します。
If Not found Then
    ' 解説: 対象が見つからないことを表示します。
    WScript.Echo "Service not found: " & serviceName
    ' 解説: 未検出を示す終了コードで終了します。
    WScript.Quit 3
' 解説: 検索結果の分岐を終了します。
End If
