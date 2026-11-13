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
`HKU`, and `HKCC`. Registry writes, deletion, import/export, binary or
multi-string access, and persistent environment-variable changes are
intentionally not included. Autologon, policy/security changes, proxy/RDP/UAC
changes, IFEO settings, and profile-hive editing are also excluded.
