> **ファイル概要:** ローカルアカウントとセキュリティの診断スクリプトの一覧。
> **実行方法:** この文書を参照し、Windowsのコマンドプロンプトでこのフォルダーから `cscript //nologo ListLocalUsers.vbs` を実行。
> **実行結果:** 例ではローカルユーザーの名前と状態がコンソールに表示されます。

# Local account and security inspection

These scripts are local, read-only inventory and status checks. Examples:

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

The HTML report contains summary counts only, is created locally, and refuses
to overwrite an existing file. WMI access and visibility vary by Windows
version and permissions.

Account/group creation or deletion, password and policy changes, group
membership changes, profile deletion, ACL changes, logoff/locking, share or
drive changes, firewall rule changes, and IPsec policy changes are
intentionally not included.

## Additional examples (items 171-200)

Each numbered item has a separate local script. State-changing requests are
implemented as status checks, preflight checks, or previews; no account,
password, membership, ACL, session, share, drive, firewall, Internet-zone, or
IPsec setting is changed.

| No. | Script | Behavior |
|---:|---|---|
| 171 | `171_LocalAccountInventory.vbs` | Lists local accounts and status fields. |
| 172 | `172_CreateLocalUserPreview.vbs` | Validates and previews a user name; creates no account. |
| 173 | `173_DeleteLocalUserPreflight.vbs` | Looks up a local account for review; does not delete it. |
| 174 | `174_LocalAccountState.vbs` | Reads enabled, lockout, and password status. |
| 175 | `175_PasswordChangePreview.vbs` | Identifies a password-change target; accepts no password. |
| 176 | `176_PasswordExpiryStatus.vbs` | Reports password expiry and required flags. |
| 177 | `177_ForcePasswordChangePreview.vbs` | Inspects password status; sets no next-logon flag. |
| 178 | `178_AdministratorsMembershipPreview.vbs` | Lists local Administrators membership without adding users. |
| 179 | `179_GroupMemberRemovalPreview.vbs` | Checks local group membership; removes no member. |
| 180 | `180_CreateLocalGroupPreview.vbs` | Validates and previews a group name; creates no group. |
| 181 | `181_AccountLockStatus.vbs` | Reads account lockout status; does not unlock accounts. |
| 182 | `182_PasswordPolicyStatus.vbs` | Reads local account password policy with `net accounts`. |
| 183 | `183_LocalSecurityPolicyStatus.vbs` | Reads audit policy; changes no local policy. |
| 184 | `184_UserProfilePathStatus.vbs` | Lists profile paths and loaded/special status. |
| 185 | `185_UserProfileRemovalPreflight.vbs` | Reports unloaded, non-special profiles last used over 90 days ago. |
| 186 | `186_ShareAclInventory.vbs` | Reads share DACL entries and access masks. |
| 187 | `187_AclGrantPreview.vbs` | Checks a local path and previews a requested access level. |
| 188 | `188_AclRemovalPreview.vbs` | Checks a local path and previews an ACL review target. |
| 189 | `189_LogoffSessionPreview.vbs` | Lists interactive sessions; logs off nobody. |
| 190 | `190_WorkstationLockPreview.vbs` | Describes the lock operation; does not lock the workstation. |
| 191 | `191_NetworkShareInventory.vbs` | Lists local share names, paths, and types. |
| 192 | `192_CreateSharePreview.vbs` | Validates a local folder and previews a share name. |
| 193 | `193_DeleteSharePreflight.vbs` | Looks up a share and backing path; deletes nothing. |
| 194 | `194_MappedDriveInventory.vbs` | Lists mapped network drives. |
| 195 | `195_DisconnectDrivePreview.vbs` | Checks a mapped drive; disconnects nothing. |
| 196 | `196_FirewallRuleAddPreview.vbs` | Validates a proposed protocol and port; adds no rule. |
| 197 | `197_FirewallRuleRemovalPreview.vbs` | Inspects a firewall rule; removes or disables nothing. |
| 198 | `198_InternetZoneStatus.vbs` | Reads the current Internet-zone level. |
| 199 | `199_IpsecPolicyStatus.vbs` | Queries connection-security rules; changes no IPsec policy. |
| 200 | `200_SecurityAuditHtmlReport.vbs` | Creates a new local HTML report containing summary counts only. |

Some queries require an elevated console. Preview scripts are not substitutes
for an approved change process; report files refuse to overwrite existing
files and reject UNC destinations.
