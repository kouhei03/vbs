# Registry and environment inspection scripts

The scripts in this directory are read-only. They can read a string or DWORD
registry value, check a key, enumerate a key's immediate subkeys, or display
only the standard process paths `SystemRoot`, `TEMP`, and `TMP`.

Examples:

```text
cscript //nologo ReadRegistryString.vbs HKCU "Control Panel\Desktop" Wallpaper
cscript //nologo ReadRegistryDword.vbs HKCU "Software\Example" Setting
cscript //nologo TestRegistryKey.vbs HKCU "Software\Example"
cscript //nologo ListRegistrySubkeys.vbs HKCU Software
cscript //nologo GetStandardEnvironmentPaths.vbs
```

Registry hives accepted by the registry scripts are `HKCU`, `HKLM`, `HKCR`,
`HKU`, and `HKCC`. Registry writes, deletion, import, repair, and persistent
environment-variable changes are intentionally not included. Autologon,
policy/security, proxy/RDP/UAC, IFEO, and profile-hive modifications are also
excluded; read-only state checks for selected items are available below.

## Additional read-only diagnostics (items 107-140)

The following scripts provide inspection or preview equivalents for the requested
items. They do not write, delete, import, or repair registry data. Export items
only check the source; environment and policy items report current state. Items
that require a state change intentionally remain out of scope.

| No. | Script | Read-only behavior |
|---:|---|---|
| 107 | `PreviewRegistryStringWrite.vbs` | Validates and previews the target and proposed string length without writing. |
| 109 | `InspectMultiStringAndBinary.vbs` | Reads REG_MULTI_SZ values or reports binary size without exposing binary data. |
| 111 | `InspectRegistryDeleteTarget.vbs` | Lists value names and types under a target key; nothing is deleted. |
| 113 | `CheckRegistryExportSource.vbs` | Checks whether the requested key is readable; does not create an export file. |
| 114 | `InspectRegFile.vbs` | Checks a REG file header and counts key sections; does not import it. |
| 115 | `ListSystemEnvironmentVariables.vbs` | Lists system environment variables without changing them. |
| 116 | `ListUserEnvironmentVariables.vbs` | Lists current-user environment variables without changing them. |
| 117 | `InspectPathDuplicates.vbs` | Reports duplicate entries in the current process PATH. |
| 118 | `InspectEnvironmentRemovalTarget.vbs` | Checks whether a user environment variable exists; does not remove it. |
| 119 | `InspectAutoLogonState.vbs` | Reports AutoAdminLogon and user name; never reads a password. |
| 120 | `InspectUSBStoragePolicy.vbs` | Reads the USBSTOR startup value without changing the policy. |
| 121 | `InspectWindowsUpdatePolicy.vbs` | Lists Windows Update policy value names and types. |
| 122 | `InspectTaskManagerPolicy.vbs` | Reads the current user's Task Manager restriction value. |
| 123 | `InspectControlPanelPolicy.vbs` | Reads the current user's Control Panel restriction value. |
| 124 | `InspectDesktopWallpaper.vbs` | Reports the current user's wallpaper path. |
| 125 | `InspectScreenSaverSettings.vbs` | Reads screen-saver activation, timeout, and secure-lock settings. |
| 126 | `InspectBrowserHomePages.vbs` | Reads the IE user homepage and Edge managed homepage values. |
| 127 | `InspectFileAssociation.vbs` | Reads the extension's default and user-selected ProgID. |
| 128 | `InspectContextMenuEntries.vbs` | Lists common shell context-menu registration subkeys. |
| 129 | `InspectApplicationFirstRunFlag.vbs` | Checks for a specified value without clearing or changing it. |
| 130 | `InspectProxySettings.vbs` | Reads proxy settings and omits values that may contain credentials. |
| 131 | `InspectRemoteDesktopPolicy.vbs` | Reads the RDP connection policy value. |
| 132 | `InspectUACSettings.vbs` | Reads selected UAC policy values without changing them. |
| 133 | `InspectExplorerDisplaySettings.vbs` | Reads hidden-file and file-extension display values. |
| 134 | `InspectIFEOBlockSettings.vbs` | Checks whether a Debugger value exists; does not reveal its command data. |
| 135 | `AuditLicenseState.vbs` | Reports licensing status without retrieving product keys. |
| 136 | `ShowAppliedGroupPolicies.vbs` | Runs the read-only `gpresult` report; does not force policy updates. |
| 137 | `CompareEnvironmentRefreshState.vbs` | Compares PATH lengths in registry and current process; sends no broadcast. |
| 138 | `InspectRegistryHiveFiles.vbs` | Reports standard hive file metadata; does not load or repair hives. |
| 139 | `InspectUserProfileHives.vbs` | Lists profile SID, loaded state, and NTUSER.DAT presence; does not mount hives. |
| 140 | `InspectRegistryAuditHistory.vbs` | Queries recent Security event 4657 entries; requires log access and auditing. |

Run scripts from an elevated console only when the requested registry location
requires it. Some diagnostics can return access-denied or unavailable results
when the account lacks permission, the key is absent, or the relevant audit
policy is not enabled. These scripts do not change those conditions.

## Additional read-only diagnostics (items 141-240)

The following grouped scripts provide safe inspection equivalents for the
requested advanced topics. They do not write registry values, change policy,
modify environment variables, load hives, or connect to remote computers.
Values that may contain credentials or identifiers are not displayed.

| No. | Script | Read-only behavior |
|---:|---|---|
| 141-160 | `141_RegistryValueAudit.vbs` | Lists value names/types and immediate subkey count for a selected local key; does not display value data. |
| 161-180 | `161_EnvironmentScopeAudit.vbs` | Compares process, user, and system environment scopes; reports PATH length, duplicate entries, and missing folders without changing them. |
| 181-205 | `181_WindowsPolicyAudit.vbs` | Reports selected Windows policy/security setting states without changing them or displaying string contents. |
| 206-240 | `206_ConfigurationInventoryAudit.vbs` | Checks selected application, runtime, browser, printer, USB, and configuration key inventories; reports counts only. |

Examples:

```text
cscript //nologo 141_RegistryValueAudit.vbs HKCU "Software"
cscript //nologo 161_EnvironmentScopeAudit.vbs
cscript //nologo 181_WindowsPolicyAudit.vbs
cscript //nologo 206_ConfigurationInventoryAudit.vbs
```

These are grouped audit examples rather than state-changing implementations.
Operations such as ACL changes, hive mounting, remote registry writes,
credential storage, policy changes, deletion, deployment, and automatic
restoration remain intentionally out of scope.
