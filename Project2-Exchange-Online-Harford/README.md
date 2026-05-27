# 📧 Project 2 — Exchange Online Administration
**Harford Property Management | Exchange Online Implementation**

---

## 📋 Project Overview

| Field | Details |
|---|---|
| **Client** | Harford Property Management |
| **Domain** | @christtech.co.uk |
| **UPN Convention** | firstname.harford@christtech.co.uk |
| **Difficulty** | 🟢 Beginner |
| **Estimated Time** | 22–28 hours |
| **Target Roles** | M365 Administrator · Exchange Administrator · IT Support Engineer |
| **Driver** | Audit approaching — no shared mailboxes, no distribution lists, spam flooding inboxes, leaver mailbox inaccessible |

---

## 🏢 Business Scenario

Harford Property Management has just migrated from on-premises Exchange to Exchange Online. No shared mailboxes were created, no distribution lists configured, spam is flooding inboxes, and nobody can access the previous IT manager's emails who left three months ago. An audit is approaching.

---

## 👥 User Naming Convention

| Field | Format | Example |
|---|---|---|
| UPN | firstname.harford@christtech.co.uk | emma.harford@christtech.co.uk |
| Display Name | Firstname Lastname | Emma Clarke |
| Alias | firstname.harford | emma.harford |

---

## 🔌 Connecting to Exchange Online

```powershell
# Install module if not already installed
Install-Module -Name ExchangeOnlineManagement -Force

# Connect using Harford admin account
Connect-ExchangeOnline -UserPrincipalName admin.harford@christtech.co.uk
```

---

## 📬 1. Shared Mailboxes

### Mailbox Architecture Design

| Mailbox | Address | Purpose | Permissions |
|---|---|---|---|
| Harford Info | info.harford@christtech.co.uk | General enquiries | Full Access + Send As → All Staff group |
| Harford Accounts | accounts.harford@christtech.co.uk | Finance communications | Full Access + Send As → Finance group |
| Harford Support | support.harford@christtech.co.uk | IT/property support | Full Access + Send As → Support group |

### PowerShell — Create Shared Mailboxes

```powershell
# Info mailbox
New-Mailbox -Name "Harford Info" -Alias "info.harford" -Shared `
  -PrimarySmtpAddress "info.harford@christtech.co.uk"

# Accounts mailbox
New-Mailbox -Name "Harford Accounts" -Alias "accounts.harford" -Shared `
  -PrimarySmtpAddress "accounts.harford@christtech.co.uk"

# Support mailbox
New-Mailbox -Name "Harford Support" -Alias "support.harford" -Shared `
  -PrimarySmtpAddress "support.harford@christtech.co.uk"
```

### PowerShell — Assign Full Access & Send As Permissions

```powershell
# Full Access — Info mailbox
Add-MailboxPermission -Identity "info.harford@christtech.co.uk" `
  -User "allstaff.harford@christtech.co.uk" `
  -AccessRights FullAccess -InheritanceType All

# Send As — Info mailbox
Add-RecipientPermission -Identity "info.harford@christtech.co.uk" `
  -Trustee "allstaff.harford@christtech.co.uk" `
  -AccessRights SendAs -Confirm:$false

# Full Access — Accounts mailbox
Add-MailboxPermission -Identity "accounts.harford@christtech.co.uk" `
  -User "finance.harford@christtech.co.uk" `
  -AccessRights FullAccess -InheritanceType All

# Send As — Accounts mailbox
Add-RecipientPermission -Identity "accounts.harford@christtech.co.uk" `
  -Trustee "finance.harford@christtech.co.uk" `
  -AccessRights SendAs -Confirm:$false

# Full Access — Support mailbox
Add-MailboxPermission -Identity "support.harford@christtech.co.uk" `
  -User "support.harford@christtech.co.uk" `
  -AccessRights FullAccess -InheritanceType All

# Send As — Support mailbox
Add-RecipientPermission -Identity "support.harford@christtech.co.uk" `
  -Trustee "support.harford@christtech.co.uk" `
  -AccessRights SendAs -Confirm:$false
```

> ⚠️ **Send As Permission Delay:** Granting Send As does not take effect immediately. It can take **up to 60 minutes** to propagate. During this window, emails sent as the shared mailbox may still display the sender's personal address. Build this waiting time into testing and set user expectations accordingly.

### Verify Shared Mailboxes

```powershell
Get-Mailbox -RecipientTypeDetails SharedMailbox | FL DisplayName, Alias, PrimarySmtpAddress
Get-MailboxPermission -Identity "info.harford@christtech.co.uk" | Where-Object {$_.IsInherited -eq $false}
Get-RecipientPermission -Identity "info.harford@christtech.co.uk" | Where-Object {$_.Trustee -ne "NT AUTHORITY\SELF"}
```

---

## 📋 2. Distribution Lists

### Distribution List Design

| List Name | Address | Members | External Senders |
|---|---|---|---|
| Harford All Staff | allstaff.harford@christtech.co.uk | All Harford employees | Restricted — must authenticate |
| Harford Finance | finance.harford@christtech.co.uk | Finance team | Restricted — must authenticate |
| Harford Management | management.harford@christtech.co.uk | Management team | Restricted — must authenticate |

### PowerShell — Create Distribution Lists

```powershell
# All Staff
New-DistributionGroup -Name "Harford All Staff" `
  -Alias "allstaff.harford" `
  -PrimarySmtpAddress "allstaff.harford@christtech.co.uk" `
  -MemberJoinRestriction Closed

# Finance
New-DistributionGroup -Name "Harford Finance" `
  -Alias "finance.harford" `
  -PrimarySmtpAddress "finance.harford@christtech.co.uk" `
  -MemberJoinRestriction Closed

# Management
New-DistributionGroup -Name "Harford Management" `
  -Alias "management.harford" `
  -PrimarySmtpAddress "management.harford@christtech.co.uk" `
  -MemberJoinRestriction Closed
```

### Restrict External Senders

```powershell
Set-DistributionGroup -Identity "allstaff.harford@christtech.co.uk" -RequireSenderAuthenticationEnabled $true
Set-DistributionGroup -Identity "finance.harford@christtech.co.uk" -RequireSenderAuthenticationEnabled $true
Set-DistributionGroup -Identity "management.harford@christtech.co.uk" -RequireSenderAuthenticationEnabled $true
```

### Add Members to Distribution Lists

```powershell
# Example — add a user to Finance list
Add-DistributionGroupMember -Identity "finance.harford@christtech.co.uk" `
  -Member "emma.harford@christtech.co.uk"
```

### Verify Distribution Lists

```powershell
Get-DistributionGroup | FL DisplayName, PrimarySmtpAddress, RequireSenderAuthenticationEnabled
Get-DistributionGroupMember -Identity "allstaff.harford@christtech.co.uk" | FL DisplayName, PrimarySmtpAddress
```

---

## 🏢 3. Room & Equipment Mailboxes

### Mailbox Design

| Mailbox | Type | Address | Capacity | Auto-Accept | Allow Conflicts |
|---|---|---|---|---|---|
| Harford Boardroom | Room | boardroom.harford@christtech.co.uk | 12 | Yes | No |
| Harford Projector | Equipment | projector.harford@christtech.co.uk | N/A | Yes | No |

### PowerShell — Create Room Mailbox

```powershell
# Create boardroom room mailbox
New-Mailbox -Name "Harford Boardroom" -Alias "boardroom.harford" -Room `
  -PrimarySmtpAddress "boardroom.harford@christtech.co.uk"

# Set capacity to 12
Set-Mailbox -Identity "boardroom.harford@christtech.co.uk" -ResourceCapacity 12

# Configure auto-accept and prevent double-bookings
Set-CalendarProcessing -Identity "boardroom.harford@christtech.co.uk" `
  -AutomateProcessing AutoAccept `
  -AllowConflicts $false `
  -MaximumDurationInMinutes 480
```

### PowerShell — Create Equipment Mailbox

```powershell
# Create projector equipment mailbox
New-Mailbox -Name "Harford Projector" -Alias "projector.harford" -Equipment `
  -PrimarySmtpAddress "projector.harford@christtech.co.uk"

# Configure auto-accept and prevent double-bookings
Set-CalendarProcessing -Identity "projector.harford@christtech.co.uk" `
  -AutomateProcessing AutoAccept `
  -AllowConflicts $false
```

### Verify Room & Equipment Mailboxes

```powershell
Get-Mailbox -Identity "boardroom.harford@christtech.co.uk" | FL Name, Alias, ResourceCapacity, RecipientTypeDetails
Get-CalendarProcessing -Identity "boardroom.harford@christtech.co.uk" | FL AutomateProcessing, AllowConflicts
Get-Mailbox -Identity "projector.harford@christtech.co.uk" | FL Name, Alias, RecipientTypeDetails
Get-CalendarProcessing -Identity "projector.harford@christtech.co.uk" | FL AutomateProcessing, AllowConflicts
```

> ⚠️ **Challenge:** The auto-accept setting alone does not prevent double-bookings. `AllowConflicts $false` must be explicitly set via PowerShell. Without it, the room mailbox will accept overlapping bookings.

---

## 📨 4. Mail Flow Rules

### Rule Design & Priority Order

| Priority | Rule Name | Scope | Action |
|---|---|---|---|
| 0 | Block Dangerous Attachments | All senders | Reject message |
| 1 | External Subject Prefix | External senders only | Prepend [EXTERNAL] |
| 2 | External Email Disclaimer | External senders only | Append disclaimer |
| 3 | HR Redirect | Sent to HR address | Redirect to HR Manager |

### Rule 1 — Block Dangerous Attachments (Priority 0)

```powershell
New-TransportRule -Name "Harford - Block Dangerous Attachments" `
  -AttachmentNameMatchesPatterns "*.exe","*.bat","*.vbs" `
  -RejectMessageReasonText "This attachment type is not permitted for security reasons." `
  -RejectMessageEnhancedStatusCode "5.7.1" `
  -Priority 0
```

### Rule 2 — [EXTERNAL] Subject Prefix (Priority 1)

```powershell
New-TransportRule -Name "Harford - External Subject Prefix" `
  -FromScope NotInOrganization `
  -PrependSubject "[EXTERNAL] " `
  -Priority 1
```

> ⚠️ **Challenge:** Scope must be `NotInOrganization` — if set incorrectly to all senders, this rule fires on internal emails too.

### Rule 3 — External Email Disclaimer (Priority 2)

```powershell
New-TransportRule -Name "Harford - External Email Disclaimer" `
  -FromScope NotInOrganization `
  -ApplyHtmlDisclaimerText "<p style='font-size:11px;color:gray;'>This email originated from outside Harford Property Management. Do not click links or open attachments unless you recognise the sender and know the content is safe.</p>" `
  -ApplyHtmlDisclaimerLocation Append `
  -ApplyHtmlDisclaimerFallbackAction Wrap `
  -Priority 2
```

### Rule 4 — HR Redirect (Priority 3)

```powershell
New-TransportRule -Name "Harford - HR Redirect" `
  -SentTo "hr.harford@christtech.co.uk" `
  -RedirectMessageTo "hrmanager.harford@christtech.co.uk" `
  -Priority 3
```

### Verify Rules & Priority Order

```powershell
Get-TransportRule | Sort-Object Priority | FL Name, Priority, State
```

> ⚠️ **Rule Ordering:** Rules are evaluated top-down by priority. Lower number = higher priority. Always test individually before applying tenant-wide.

---

## 🛡️ 5. Anti-Spam Policy

### Policy Design

| Setting | Default Policy | Harford Management Policy |
|---|---|---|
| Bulk Email Threshold | 7 | 5 (stricter) |
| Scope | All users | management.harford@christtech.co.uk only |
| Spam Action | Move to Junk | Move to Junk |
| High Confidence Spam | Quarantine | Quarantine |

### PowerShell — Create Anti-Spam Policy for Management

```powershell
# Create the policy
New-HostedContentFilterPolicy -Name "Harford Management Anti-Spam" `
  -BulkThreshold 5 `
  -SpamAction MoveToJmf `
  -HighConfidenceSpamAction Quarantine

# Scope it to Harford Management group only
New-HostedContentFilterRule -Name "Harford Management Anti-Spam Rule" `
  -HostedContentFilterPolicy "Harford Management Anti-Spam" `
  -SentToMemberOf "management.harford@christtech.co.uk" `
  -Priority 0
```

### Verify Anti-Spam Policy

```powershell
Get-HostedContentFilterPolicy -Identity "Harford Management Anti-Spam" | FL Name, BulkThreshold
Get-HostedContentFilterRule -Identity "Harford Management Anti-Spam Rule" | FL Name, SentToMemberOf, Priority
```

---

## 🔐 6. Modern Authentication — Block Basic Auth

### Policy Design

| Setting | Value |
|---|---|
| Protocol | Exchange ActiveSync |
| Basic Authentication | Blocked |
| Modern Authentication | Required |
| Scope | Entire Harford tenant |

### PowerShell — Block Basic Authentication

```powershell
# Create authentication policy
New-AuthenticationPolicy -Name "Harford - Block Basic Auth" `
  -AllowBasicAuthActiveSync $false `
  -AllowBasicAuthImap $false `
  -AllowBasicAuthPop $false `
  -AllowBasicAuthSmtp $false `
  -AllowBasicAuthWebServices $false

# Apply to all Harford users
Get-Mailbox -ResultSize Unlimited | ForEach-Object {
    Set-User -Identity $_.Identity -AuthenticationPolicy "Harford - Block Basic Auth"
}
```

### Verify

```powershell
Get-AuthenticationPolicy -Identity "Harford - Block Basic Auth" | FL
Get-User -ResultSize Unlimited | Where-Object {
    $_.AuthenticationPolicy -eq "Harford - Block Basic Auth"
} | FL DisplayName, UserPrincipalName, AuthenticationPolicy
```

---

## 🔓 7. Leaver Mailbox Access

### Legal & HR Approval Process

> ⚠️ **GDPR Requirement:** Access to a leaver's mailbox must NOT be granted without documented authorisation.

| Step | Action | Owner |
|---|---|---|
| 1 | Submit formal mailbox access request | IT Manager |
| 2 | HR Manager provides written approval | HR Manager |
| 3 | Legal/DPO confirms GDPR compliance | Data Protection Officer |
| 4 | Approval documented and filed | IT Manager |
| 5 | IT Admin granted access | IT Administrator |
| 6 | Access removed when no longer needed | IT Manager |

### PowerShell — Grant Full Access to Leaver Mailbox

```powershell
# Confirm leaver mailbox exists
Get-Mailbox -Identity "formeritmanager.harford@christtech.co.uk" | FL DisplayName, PrimarySmtpAddress

# Grant Full Access to IT Admin
Add-MailboxPermission -Identity "formeritmanager.harford@christtech.co.uk" `
  -User "itadmin.harford@christtech.co.uk" `
  -AccessRights FullAccess `
  -InheritanceType All

# Verify
Get-MailboxPermission -Identity "formeritmanager.harford@christtech.co.uk" | `
  Where-Object {$_.User -like "*itadmin.harford*"} | FL User, AccessRights
```

### Remove Access When No Longer Required

```powershell
Remove-MailboxPermission -Identity "formeritmanager.harford@christtech.co.uk" `
  -User "itadmin.harford@christtech.co.uk" `
  -AccessRights FullAccess `
  -Confirm:$false
```

> **GDPR Considerations:**
> - Access must be **purpose-limited** — legitimate business continuity only
> - Access must be **time-limited** — remove as soon as purpose is fulfilled
> - All access must be **logged** — who accessed, when, and why
> - Unauthorised access may constitute a data breach under UK GDPR

---

## 📚 Runbooks

### Runbook 1 — Create Shared Mailbox

| Step | Action |
|---|---|
| 1 | Connect: `Connect-ExchangeOnline -UserPrincipalName admin.harford@christtech.co.uk` |
| 2 | Run `New-Mailbox` with `-Shared` flag and address format `function.harford@christtech.co.uk` |
| 3 | Assign Full Access via `Add-MailboxPermission` |
| 4 | Assign Send As via `Add-RecipientPermission` |
| 5 | Wait up to **60 minutes** for Send As to propagate |
| 6 | Test by sending as the shared mailbox |
| 7 | Confirm sender address shows shared mailbox address |

### Runbook 2 — Create Distribution List

| Step | Action |
|---|---|
| 1 | Run `New-DistributionGroup` with address format `groupname.harford@christtech.co.uk` |
| 2 | Set `RequireSenderAuthenticationEnabled $true` |
| 3 | Add members via `Add-DistributionGroupMember` |
| 4 | Test by sending from an external address — should be rejected |
| 5 | Verify with `Get-DistributionGroup` |

### Runbook 3 — Grant Mailbox Permissions

| Step | Action |
|---|---|
| 1 | Confirm target mailbox exists |
| 2 | Run `Add-MailboxPermission` for Full Access |
| 3 | Run `Add-RecipientPermission` for Send As |
| 4 | Note: Send As may take up to **60 minutes** to propagate |
| 5 | Verify with `Get-MailboxPermission` and `Get-RecipientPermission` |

### Runbook 4 — Create Room Mailbox

| Step | Action |
|---|---|
| 1 | Run `New-Mailbox` with `-Room` flag and address format `roomname.harford@christtech.co.uk` |
| 2 | Set capacity with `Set-Mailbox -ResourceCapacity 12` |
| 3 | Set `AutomateProcessing AutoAccept` via `Set-CalendarProcessing` |
| 4 | Set `AllowConflicts $false` to prevent double-bookings |
| 5 | Verify with `Get-CalendarProcessing` |

### Runbook 5 — Access Leaver Mailbox

| Step | Action |
|---|---|
| 1 | Obtain written approval from HR Manager |
| 2 | Obtain GDPR sign-off from Data Protection Officer |
| 3 | File approval documentation before proceeding |
| 4 | Run `Add-MailboxPermission` to grant Full Access |
| 5 | Log: who accessed, when, and why |
| 6 | Remove access via `Remove-MailboxPermission` when done |

### Runbook 6 — Create Transport Rule

| Step | Action |
|---|---|
| 1 | Plan rule conditions, actions, scope, and priority |
| 2 | Run `New-TransportRule` with all parameters |
| 3 | Test on a small scope before applying tenant-wide |
| 4 | Verify with `Get-TransportRule | Sort-Object Priority` |
| 5 | Monitor message trace for unexpected rule matches |

---

## 📋 Change Record

| Field | Details |
|---|---|
| **Change ID** | CHG-002 |
| **Title** | Exchange Online Initial Build — Harford Property Management |
| **Date** | 27/05/2026 |
| **Implemented By** | itadmin.harford@christtech.co.uk |
| **Approved By** | IT Manager |
| **Description** | Created 3 shared mailboxes, 3 distribution lists, 1 room mailbox, 1 equipment mailbox, 4 mail flow rules, anti-spam policy for management, blocked basic auth, granted leaver mailbox access with GDPR approval |
| **Status** | ✅ Completed |
| **Rollback Plan** | Remove mail flow rules, delete shared mailboxes, revert auth policy, remove leaver mailbox permissions |

---

## 🔧 Troubleshooting Guide

### Issue 1: Send As Still Showing Personal Address
**Cause:** Send As propagation delay — up to 60 minutes.
**Fix:** Wait and verify:
```powershell
Get-RecipientPermission -Identity "info.harford@christtech.co.uk" | Where-Object {$_.Trustee -ne "NT AUTHORITY\SELF"}
```

### Issue 2: [EXTERNAL] Prefix Firing on Internal Emails
**Cause:** Rule scope set to all senders instead of external only.
**Fix:**
```powershell
Set-TransportRule -Identity "Harford - External Subject Prefix" -FromScope NotInOrganization
```

### Issue 3: Room Mailbox Accepting Double-Bookings
**Cause:** `AllowConflicts` not explicitly set to `$false`.
**Fix:**
```powershell
Set-CalendarProcessing -Identity "boardroom.harford@christtech.co.uk" -AllowConflicts $false
```

### Issue 4: External Senders Reaching Distribution List
**Cause:** `RequireSenderAuthenticationEnabled` not set.
**Fix:**
```powershell
Set-DistributionGroup -Identity "allstaff.harford@christtech.co.uk" -RequireSenderAuthenticationEnabled $true
```

### Issue 5: Leaver Mailbox Access Denied
**Cause:** Permission not granted or not yet propagated.
**Fix:**
```powershell
Add-MailboxPermission -Identity "formeritmanager.harford@christtech.co.uk" `
  -User "itadmin.harford@christtech.co.uk" `
  -AccessRights FullAccess -InheritanceType All
```

---

## ⚡ Challenges Faced

| Challenge | Resolution |
|---|---|
| Send As permission delay | Documented 60-minute delay in runbook; communicated to users before testing |
| Mail flow rule ordering conflicts | Assigned explicit priorities 0–3; tested each rule individually |
| Room mailbox accepting double-bookings | Set `AllowConflicts $false` explicitly — auto-accept alone is not sufficient |
| [EXTERNAL] prefix firing on internal emails | Corrected scope to `NotInOrganization` |
| Leaver mailbox GDPR concerns | Added formal HR and DPO approval step to runbook before any technical access |

---

## 💡 Lessons Learned

1. **Always test mail flow rules in a small scope before applying tenant-wide** — one incorrectly scoped rule affects every email in the organisation.
2. **Permission propagation delays are real** — build waiting time into runbooks and set user expectations.
3. **Room mailboxes require more configuration than expected** — auto-accept alone is insufficient.
4. **Documenting leaver mailbox access is as important as the technical steps** — the HR and legal approval step must exist before anyone needs to use it.
5. **The [EXTERNAL] prefix is one of the most effective low-cost security controls** — document the rationale so users understand why their emails look different.

---

## 📁 Repository Structure

```
Greenwood/
├── Project2-Exchange-Online-Harford/
│   ├── README.md
│   ├── Design-Documents/
│   │   ├── Mailbox-Architecture.md
│   │   ├── Shared-Mailbox-Design.md
│   │   ├── Mail-Flow-Rules.md
│   │   ├── Anti-Spam-Configuration.md
│   │   └── Mobile-Access-Policy.md
│   ├── Runbooks/
│   │   ├── Create-Shared-Mailbox.md
│   │   ├── Create-Distribution-List.md
│   │   ├── Grant-Mailbox-Permissions.md
│   │   ├── Create-Room-Mailbox.md
│   │   ├── Access-Leaver-Mailbox.md
│   │   └── Create-Transport-Rule.md
│   ├── Change-Record/
│   │   └── Change-Record-Exchange-Online.md
│   └── Troubleshooting-Guide/
│       └── Troubleshooting-Guide.md
```

---

*Harford Property Management — Exchange Online Administration | Project 2 | May 2026*
