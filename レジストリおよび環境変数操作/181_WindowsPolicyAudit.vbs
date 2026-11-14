Option Explicit

Dim registry, checks, fields, hiveName, hive, subKey, valueName
Dim dataType, result, data, index, displayValue

Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

' Each entry is hive|key|value|DWORD-or-STRING. Only selected policy state is reported.
checks = Array( _
    "HKCU|Software\Policies\Microsoft\Windows\Explorer|DisableNotificationCenter|DWORD", _
    "HKLM|SOFTWARE\Policies\Microsoft\Windows\Personalization|NoLockScreen|DWORD", _
    "HKLM|SOFTWARE\Policies\Microsoft\Windows\AppPrivacy|LetAppsRunInBackground|DWORD", _
    "HKLM|SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer|NoDriveTypeAutoRun|DWORD", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Policies\Explorer|NoSaveSettings|DWORD", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Policies\Explorer|NoSetFolders|DWORD", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowCpl||KEY", _
    "HKLM|SOFTWARE\Policies\Microsoft\Edge|HideFirstRunExperience|DWORD", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Themes\Personalize|AppsUseLightTheme|DWORD", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Themes\Personalize|SystemUsesLightTheme|DWORD", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Search|SearchboxTaskbarMode|DWORD", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState|FullPath|DWORD", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects|VisualFXSetting|DWORD", _
    "HKLM|SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon|AutoAdminLogon|STRING", _
    "HKCU|Control Panel\Desktop|ScreenSaveActive|STRING", _
    "HKLM|SYSTEM\CurrentControlSet\Control\Terminal Server|fDenyTSConnections|DWORD", _
    "HKLM|SOFTWARE\Policies\Microsoft\Windows Defender\Spynet|SpynetReporting|DWORD", _
    "HKLM|SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile|DisableNotifications|DWORD", _
    "HKLM|SOFTWARE\Policies\Microsoft\WindowsFirewall\StandardProfile|DisableNotifications|DWORD", _
    "HKCU|Software\Policies\Microsoft\Windows\WindowsUpdate|DisableWindowsUpdateAccess|DWORD", _
    "HKLM|SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU|NoAutoRebootWithLoggedOnUsers|DWORD", _
    "HKLM|SYSTEM\CurrentControlSet\Control\Lsa|EveryoneIncludesAnonymous|DWORD", _
    "HKLM|SYSTEM\CurrentControlSet\Control\Power|HiberbootEnabled|DWORD", _
    "HKLM|SOFTWARE\Policies\Microsoft\Windows\System|EnableSmartScreen|DWORD", _
    "HKCU|Software\Microsoft\Windows\CurrentVersion\Policies\System|DisableTaskMgr|DWORD" _
)

For index = 0 To UBound(checks)
    fields = Split(checks(index), "|")
    hiveName = fields(0)
    subKey = fields(1)
    valueName = fields(2)
    dataType = fields(3)

    Select Case hiveName
        Case "HKCU": hive = &H80000001
        Case "HKLM": hive = &H80000002
        Case Else
            WScript.Echo "Internal error: unsupported hive " & hiveName
            WScript.Quit 1
    End Select

    If dataType = "KEY" Then
        result = registry.EnumKey(hive, subKey, data)
    ElseIf dataType = "DWORD" Then
        result = registry.GetDWORDValue(hive, subKey, valueName, data)
    Else
        result = registry.GetStringValue(hive, subKey, valueName, data)
    End If

    displayValue = "unavailable (WMI return code " & result & ")"
    If result = 0 Then
        If dataType = "KEY" Then
            If IsArray(data) Then
                displayValue = "configured entries=" & (UBound(data) + 1)
            Else
                displayValue = "configured entries=0"
            End If
        ElseIf dataType = "STRING" Then
            If Len(CStr(data)) = 0 Then
                displayValue = "present (value omitted)"
            Else
                displayValue = "present (value omitted; length=" & Len(CStr(data)) & ")"
            End If
        Else
            displayValue = CStr(data)
        End If
    ElseIf result = 2 Then
        displayValue = "not configured"
    End If

    WScript.Echo hiveName & "\" & subKey & "\" & valueName & " = " & displayValue
Next

WScript.Echo "Audit only: policy values were not changed."
