' ファイル概要: アプリケーションや周辺機器のレジストリキーの数を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 206_ConfigurationInventoryAudit.vbs
' 実行結果: アプリケーションや周辺機器のレジストリキーの数を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim registry, targets, fields, hiveName, hive, subKey
Dim result, subKeys, valueNames, valueTypes, subkeyCount, valueCount, index

Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

' The inventory reports only key presence and counts, not device IDs or stored values.
targets = Array( _
    "HKCU|Software\Company\ExampleApp", _
    "HKLM|SOFTWARE\Company\ExampleApp\Licensing", _
    "HKLM|SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", _
    "HKCU|Software\Company\Deployment", _
    "HKLM|SOFTWARE\Company\RegistryAudit", _
    "HKCU|Software\Company\RegistrySnapshots", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU", _
    "HKLM|SOFTWARE\Company\ConfigurationSource", _
    "HKLM|SOFTWARE\Company\RegistryChangeAudit", _
    "HKLM|SOFTWARE\Company\RegistryRecovery", _
    "HKLM|SOFTWARE\Policies\Microsoft\Windows", _
    "HKLM|SOFTWARE\Microsoft\InetStp", _
    "HKCR|CLSID", _
    "HKLM|SYSTEM\CurrentControlSet\Services\LanmanServer\Shares", _
    "HKCU|Software\Microsoft\Windows NT\CurrentVersion\Virtualization", _
    "HKLM|SYSTEM\CurrentControlSet\Control\Print\Printers", _
    "HKLM|SYSTEM\CurrentControlSet\Control\Network", _
    "HKCU|Software\Company\Maintenance", _
    "HKLM|SOFTWARE\Microsoft\NET Framework Setup\NDP", _
    "HKLM|SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist", _
    "HKLM|SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Internet Settings", _
    "HKLM|SYSTEM\CurrentControlSet\Control\TimeZoneInformation", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs", _
    "HKLM|SYSTEM\CurrentControlSet\Enum\USBSTOR", _
    "HKLM|SYSTEM\CurrentControlSet\Control\Print\Monitors", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Themes", _
    "HKLM|SOFTWARE\Company\RegistryOperationLog", _
    "HKLM|SYSTEM\CurrentControlSet\Control\Session Manager\Environment", _
    "HKCU|Environment", _
    "HKLM|SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", _
    "HKLM|SOFTWARE\Microsoft\Windows NT\CurrentVersion", _
    "HKLM|SOFTWARE\Microsoft\Windows\CurrentVersion\Policies", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Policies", _
    "HKLM|SYSTEM\CurrentControlSet\Control\Print\Printers" _
)

For index = 0 To UBound(targets)
    fields = Split(targets(index), "|")
    hiveName = fields(0)
    subKey = fields(1)

    Select Case hiveName
        Case "HKCR": hive = &H80000000
        Case "HKCU": hive = &H80000001
        Case "HKLM": hive = &H80000002
        Case "HKU": hive = &H80000003
        Case "HKCC": hive = &H80000005
        Case Else
            WScript.Echo "Internal error: unsupported hive " & hiveName
            WScript.Quit 1
    End Select

    result = registry.EnumKey(hive, subKey, subKeys)
    If result = 0 Then
        subkeyCount = 0
        If IsArray(subKeys) Then subkeyCount = UBound(subKeys) + 1

        result = registry.EnumValues(hive, subKey, valueNames, valueTypes)
        valueCount = 0
        If result = 0 And IsArray(valueNames) Then valueCount = UBound(valueNames) + 1
        If result <> 0 Then
            WScript.Echo hiveName & "\" & subKey & " present; values unavailable (WMI return code " & result & ")"
        Else
            WScript.Echo hiveName & "\" & subKey & " present; subkeys=" & subkeyCount & "; values=" & valueCount
        End If
    ElseIf result = 2 Then
        WScript.Echo hiveName & "\" & subKey & " not present"
    Else
        WScript.Echo hiveName & "\" & subKey & " unavailable (WMI return code " & result & ")"
    End If
Next

WScript.Echo "Audit only: no registry data was modified, deleted, exported, or restored."
