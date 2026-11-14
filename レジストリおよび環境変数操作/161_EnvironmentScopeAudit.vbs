' ファイル概要: 環境変数スコープとPATHの重複・欠落を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 161_EnvironmentScopeAudit.vbs
' 実行結果: 環境変数スコープとPATHの重複・欠落を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim shell, fso, scopeName, environment, entry, separator, variableName
Dim variableValue, pathValue, pathItems, item, normalized, seen, missingCount
Dim duplicateCount, pathCount, totalCount, scopeIndex, currentScope, scopeMaps(2)

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

For Each scopeName In Array("PROCESS", "USER", "SYSTEM")
    Set environment = shell.Environment(scopeName)
    Select Case scopeName
        Case "PROCESS": scopeIndex = 0
        Case "USER": scopeIndex = 1
        Case "SYSTEM": scopeIndex = 2
    End Select
    Set currentScope = CreateObject("Scripting.Dictionary")
    ' [処理解説] 辞書のキーは大文字・小文字を区別せず照合します。
    currentScope.CompareMode = vbTextCompare
    Set scopeMaps(scopeIndex) = currentScope
    totalCount = 0
    pathValue = ""

    WScript.Echo "[" & scopeName & "]"
    For Each entry In environment
        separator = InStr(1, entry, "=", vbBinaryCompare)
        If separator > 0 Then
            variableName = Left(entry, separator - 1)
            variableValue = Mid(entry, separator + 1)
            totalCount = totalCount + 1
            If Not currentScope.Exists(variableName) Then currentScope.Add variableName, True
            If UCase(variableName) = "PATH" Then pathValue = variableValue
            WScript.Echo variableName & " (length=" & Len(variableValue) & ")"
        End If
    Next

    WScript.Echo "Variable count: " & totalCount
    If Len(pathValue) > 0 Then
        Set seen = CreateObject("Scripting.Dictionary")
        seen.CompareMode = vbTextCompare
        pathItems = Split(pathValue, ";")
        pathCount = 0
        missingCount = 0
        duplicateCount = 0

        For Each item In pathItems
            item = Trim(item)
            If Len(item) > 0 Then
                pathCount = pathCount + 1
                normalized = shell.ExpandEnvironmentStrings(item)
                If Right(normalized, 1) = "\" Then
                    normalized = Left(normalized, Len(normalized) - 1)
                End If
                If seen.Exists(normalized) Then
                    duplicateCount = duplicateCount + 1
                Else
                    seen.Add normalized, True
                    If InStr(normalized, "%") = 0 Then
                        If Not fso.FolderExists(normalized) Then missingCount = missingCount + 1
                    End If
                End If
            End If
        Next

        WScript.Echo "PATH length: " & Len(pathValue)
        WScript.Echo "PATH entries: " & pathCount
        WScript.Echo "Duplicate PATH entries: " & duplicateCount
        WScript.Echo "Missing PATH folders (unexpanded entries excluded): " & missingCount
    Else
        WScript.Echo "PATH is not set in this scope."
    End If
Next

For Each variableName In scopeMaps(1).Keys
    If scopeMaps(2).Exists(variableName) Then
        WScript.Echo "Defined in both USER and SYSTEM scopes: " & variableName
    End If
Next

WScript.Echo "Environment values are not printed or changed."
