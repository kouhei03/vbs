# Local account and security inspection
<!-- 解説: この見出しは「Local account and security inspection」を説明する節のタイトルです。 -->

These scripts are local, read-only inventory and status checks. Examples:
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: These scripts are local, read-only inventory and status checks. Examples: -->

```text
cscript //nologo ListLocalUsers.vbs
cscript //nologo GetLocalUserStatus.vbs alice
cscript //nologo GetPasswordPolicy.vbs
cscript //nologo ListUserProfiles.vbs
cscript //nologo ListNetworkShares.vbs
cscript //nologo GetShareAcl.vbs Public
cscript //nologo ListMappedDrives.vbs
cscript //nologo GetFirewallProfileStatus.vbs
cscript //nologo CreateSecurityAuditSummary.vbs C:\Temp\security-summary.html
```
<!-- 解説: この行はREADME内の実行例コードブロックの開始を示します。 -->
<!-- 解説: このコマンド例はListLocalUsers.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はGetLocalUserStatus.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はGetPasswordPolicy.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はListUserProfiles.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はListNetworkShares.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はGetShareAcl.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はListMappedDrives.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はGetFirewallProfileStatus.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: このコマンド例はCreateSecurityAuditSummary.vbsをWindows Script Hostのコンソールモードで実行します。 -->
<!-- 解説: この行は実行例コードブロックの終端です。 -->

The HTML report contains summary counts only, is created locally, and refuses
to overwrite an existing file. WMI access and visibility vary by Windows
version and permissions.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: The HTML report contains summary counts only, is created locally, and refuses -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: to overwrite an existing file. WMI access and visibility vary by Windows -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: version and permissions. -->

Account/group creation or deletion, password and policy changes, group
membership changes, profile deletion, ACL changes, logoff/locking, share or
drive changes, firewall rule changes, and IPsec policy changes are
intentionally not included.
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: Account/group creation or deletion, password and policy changes, group -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: membership changes, profile deletion, ACL changes, logoff/locking, share or -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: drive changes, firewall rule changes, and IPsec policy changes are -->
<!-- 解説: この本文行はREADMEの使い方、対象範囲、または注意事項を説明しています: intentionally not included. -->
