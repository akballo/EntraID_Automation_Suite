# 🚀 Microsoft 365 Administration Portfolio
### PowerShell Projects — Proof of 5+ Years M365 Experience
> **Author:** Christian (Atsou Komi Ballo) | M365 Administrator & Engineer | Philippines
> **Skills:** Entra ID · Exchange Online · Teams · Intune · SharePoint · Conditional Access · DLP · Purview

---

## 📁 Portfolio Overview

This portfolio contains **5 real-world PowerShell projects** that demonstrate hands-on Microsoft 365 administration skills across identity management, governance, security, compliance, and device management.

| # | Project | Key Skills |
|---|---------|------------|
| 1 | [Entra ID Automation Suite](#project-1) | Entra ID, Conditional Access, PowerShell |
| 2 | [Teams Governance Dashboard](#project-2) | Microsoft Graph API, Teams Admin |
| 3 | [Exchange Online Management Toolkit](#project-3) | Exchange Online, Mailbox Management |
| 4 | [DLP & Compliance Auditing Tool](#project-4) | Purview, DLP, Compliance Reports |
| 5 | [Intune Device Management Dashboard](#project-5) | Intune, MDM, Security Posture |

---

## PROJECT 1
## 🔐 Entra ID Automation Suite {#project-1}

### 📌 Problem It Solves
Manually creating users, assigning licenses, and configuring Conditional Access policies is time-consuming and error-prone at scale. This suite automates the entire identity lifecycle.

### 🛠️ Technologies Used
- Microsoft Graph PowerShell SDK
- Entra ID (Azure AD)
- Conditional Access Policies
- Microsoft 365 License Management

### 📂 Repository Structure
```
EntraID-Automation-Suite/
├── README.md
├── scripts/
│   ├── 01-BulkCreateUsers.ps1
│   ├── 02-AssignLicensesByDepartment.ps1
│   ├── 03-ConditionalAccessSetup.ps1
│   └── 04-DisableInactiveUsers.ps1
├── csv-templates/
│   └── users-template.csv
└── docs/
    └── setup-guide.md
```

### ⚙️ Prerequisites
```powershell
# Step 1: Install required modules
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module AzureAD -Scope CurrentUser -Force

# Step 2: Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "Policy.ReadWrite.ConditionalAccess"
```

### 📄 Script 1 — Bulk Create Users from CSV
```powershell
# 01-BulkCreateUsers.ps1
# Description: Creates multiple users in Entra ID from a CSV file
# CSV Headers Required: DisplayName, UserPrincipalName, Department, JobTitle, Password

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath
)

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All"

# Import CSV file
$Users = Import-Csv -Path $CsvPath

$Results = @()

foreach ($User in $Users) {
    try {
        # Build password profile
        $PasswordProfile = @{
            Password = $User.Password
            ForceChangePasswordNextSignIn = $true
        }

        # Create user
        $NewUser = New-MgUser `
            -DisplayName $User.DisplayName `
            -UserPrincipalName $User.UserPrincipalName `
            -Department $User.Department `
            -JobTitle $User.JobTitle `
            -PasswordProfile $PasswordProfile `
            -AccountEnabled $true `
            -MailNickname ($User.UserPrincipalName.Split("@")[0])

        Write-Host "✅ Created: $($User.DisplayName)" -ForegroundColor Green

        $Results += [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            UPN               = $User.UserPrincipalName
            Department        = $User.Department
            Status            = "Success"
            UserId            = $NewUser.Id
        }
    }
    catch {
        Write-Host "❌ Failed: $($User.DisplayName) — $($_.Exception.Message)" -ForegroundColor Red
        $Results += [PSCustomObject]@{
            DisplayName = $User.DisplayName
            UPN         = $User.UserPrincipalName
            Department  = $User.Department
            Status      = "Failed: $($_.Exception.Message)"
            UserId      = "N/A"
        }
    }
}

# Export results report
$Results | Export-Csv -Path ".\BulkCreateUsers-Report.csv" -NoTypeInformation
Write-Host "`n📊 Report saved to BulkCreateUsers-Report.csv" -ForegroundColor Cyan
```

### 📄 Script 2 — Assign Licenses by Department
```powershell
# 02-AssignLicensesByDepartment.ps1
# Description: Assigns M365 licenses to users based on their department

Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All"

# Define department-to-license mapping
# Get SKU IDs by running: Get-MgSubscribedSku | Select SkuPartNumber, SkuId
$LicenseMap = @{
    "IT"        = "YOUR-E5-SKU-ID-HERE"
    "Sales"     = "YOUR-E3-SKU-ID-HERE"
    "HR"        = "YOUR-BUSINESS-PREMIUM-SKU-ID-HERE"
    "Finance"   = "YOUR-E3-SKU-ID-HERE"
}

# Get all users grouped by department
foreach ($Department in $LicenseMap.Keys) {
    $Users = Get-MgUser -Filter "department eq '$Department'" -All

    foreach ($User in $Users) {
        try {
            $LicenseId = $LicenseMap[$Department]

            Set-MgUserLicense -UserId $User.Id `
                -AddLicenses @{SkuId = $LicenseId} `
                -RemoveLicenses @()

            Write-Host "✅ Licensed [$Department]: $($User.DisplayName)" -ForegroundColor Green
        }
        catch {
            Write-Host "❌ Failed [$Department]: $($User.DisplayName)" -ForegroundColor Red
        }
    }
}
```

### 📄 Script 3 — Conditional Access Policy Setup
```powershell
# 03-ConditionalAccessSetup.ps1
# Description: Creates a Conditional Access policy requiring MFA for all users
# except break-glass admin accounts

Connect-MgGraph -Scopes "Policy.ReadWrite.ConditionalAccess", "Directory.ReadWrite.All"

# Define break-glass accounts to exclude (use Object IDs)
$ExcludedUserIds = @(
    "BREAKGLASS-ACCOUNT-OBJECT-ID-1",
    "BREAKGLASS-ACCOUNT-OBJECT-ID-2"
)

$PolicyParams = @{
    DisplayName = "CORP-CA-001: Require MFA for All Users"
    State       = "enabledForReportingButNotEnforced"  # Use 'enabled' when ready for production
    Conditions  = @{
        Users = @{
            IncludeUsers = @("All")
            ExcludeUsers = $ExcludedUserIds
        }
        Applications = @{
            IncludeApplications = @("All")
        }
        ClientAppTypes = @("all")
    }
    GrantControls = @{
        Operator         = "OR"
        BuiltInControls  = @("mfa")
    }
}

try {
    $Policy = New-MgIdentityConditionalAccessPolicy -BodyParameter $PolicyParams
    Write-Host "✅ Conditional Access Policy Created: $($Policy.DisplayName)" -ForegroundColor Green
    Write-Host "   Policy ID: $($Policy.Id)" -ForegroundColor Cyan
}
catch {
    Write-Host "❌ Failed to create policy: $($_.Exception.Message)" -ForegroundColor Red
}
```

### 📄 Script 4 — Disable Inactive Users
```powershell
# 04-DisableInactiveUsers.ps1
# Description: Finds users who haven't signed in for 90+ days and disables their accounts

Connect-MgGraph -Scopes "User.ReadWrite.All", "AuditLog.Read.All"

$InactiveDays = 90
$CutoffDate = (Get-Date).AddDays(-$InactiveDays).ToString("yyyy-MM-ddTHH:mm:ssZ")

# Get users with sign-in activity
$InactiveUsers = Get-MgUser -All `
    -Property "Id,DisplayName,UserPrincipalName,SignInActivity,AccountEnabled,Department" `
    -Filter "signInActivity/lastSignInDateTime le $CutoffDate and accountEnabled eq true"

Write-Host "Found $($InactiveUsers.Count) inactive users (90+ days)" -ForegroundColor Yellow

$Results = @()

foreach ($User in $InactiveUsers) {
    try {
        # Disable the account
        Update-MgUser -UserId $User.Id -AccountEnabled $false

        Write-Host "🔒 Disabled: $($User.DisplayName)" -ForegroundColor Yellow

        $Results += [PSCustomObject]@{
            DisplayName    = $User.DisplayName
            UPN            = $User.UserPrincipalName
            Department     = $User.Department
            LastSignIn     = $User.SignInActivity.LastSignInDateTime
            Action         = "Disabled"
        }
    }
    catch {
        Write-Host "❌ Failed: $($User.DisplayName)" -ForegroundColor Red
    }
}

$Results | Export-Csv -Path ".\InactiveUsers-Report.csv" -NoTypeInformation
Write-Host "`n📊 Report saved to InactiveUsers-Report.csv" -ForegroundColor Cyan
```

---

## PROJECT 2
## 📊 Teams Governance Dashboard {#project-2}

### 📌 Problem It Solves
Organizations often lose track of orphaned Teams, excessive guest users, and inactive channels. This toolkit audits your entire Teams environment and generates a compliance report.

### 🛠️ Technologies Used
- Microsoft Graph PowerShell SDK
- Microsoft Teams Admin
- Teams Governance & Compliance

### 📂 Repository Structure
```
Teams-Governance-Dashboard/
├── README.md
├── scripts/
│   ├── 01-GetAllTeams.ps1
│   ├── 02-AuditGuestUsers.ps1
│   ├── 03-FindOrphanedTeams.ps1
│   └── 04-GenerateGovernanceReport.ps1
└── reports/
    └── (generated reports appear here)
```

### ⚙️ Prerequisites
```powershell
# Install required modules
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module MicrosoftTeams -Scope CurrentUser -Force

# Connect
Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All", "TeamMember.Read.All"
Connect-MicrosoftTeams
```

### 📄 Script 1 — Get All Teams Inventory
```powershell
# 01-GetAllTeams.ps1
# Description: Gets a full inventory of all Teams in the tenant

Connect-MgGraph -Scopes "Group.Read.All", "TeamMember.Read.All"

$AllTeams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All `
    -Property "Id,DisplayName,Description,CreatedDateTime,Visibility,Mail"

$TeamInventory = @()

foreach ($Team in $AllTeams) {
    # Get member count
    $Members = Get-MgGroupMember -GroupId $Team.Id -All
    $Owners  = Get-MgGroupOwner  -GroupId $Team.Id -All

    $TeamInventory += [PSCustomObject]@{
        TeamName        = $Team.DisplayName
        TeamId          = $Team.Id
        Visibility      = $Team.Visibility
        MemberCount     = $Members.Count
        OwnerCount      = $Owners.Count
        CreatedDate     = $Team.CreatedDateTime
        Email           = $Team.Mail
    }
}

$TeamInventory | Export-Csv -Path ".\reports\Teams-Inventory.csv" -NoTypeInformation
Write-Host "✅ Teams inventory exported: $($TeamInventory.Count) teams found" -ForegroundColor Green
```

### 📄 Script 2 — Audit Guest Users in Teams
```powershell
# 02-AuditGuestUsers.ps1
# Description: Reports all guest users across all Teams

Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All"

$AllTeams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All
$GuestReport = @()

foreach ($Team in $AllTeams) {
    $Members = Get-MgGroupMember -GroupId $Team.Id -All

    foreach ($Member in $Members) {
        $User = Get-MgUser -UserId $Member.Id -Property "DisplayName,UserPrincipalName,UserType,Mail" -ErrorAction SilentlyContinue

        if ($User.UserType -eq "Guest") {
            $GuestReport += [PSCustomObject]@{
                TeamName          = $Team.DisplayName
                GuestName         = $User.DisplayName
                GuestEmail        = $User.Mail
                GuestUPN          = $User.UserPrincipalName
            }
        }
    }
}

$GuestReport | Export-Csv -Path ".\reports\Teams-GuestUsers.csv" -NoTypeInformation
Write-Host "✅ Found $($GuestReport.Count) guest users across all Teams" -ForegroundColor Yellow
```

### 📄 Script 3 — Find Orphaned Teams (No Owners)
```powershell
# 03-FindOrphanedTeams.ps1
# Description: Identifies Teams that have no owners (orphaned)

Connect-MgGraph -Scopes "Group.Read.All"

$AllTeams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All
$OrphanedTeams = @()

foreach ($Team in $AllTeams) {
    $Owners = Get-MgGroupOwner -GroupId $Team.Id -All

    if ($Owners.Count -eq 0) {
        $OrphanedTeams += [PSCustomObject]@{
            TeamName    = $Team.DisplayName
            TeamId      = $Team.Id
            CreatedDate = $Team.CreatedDateTime
            Status      = "ORPHANED - No Owners"
        }
        Write-Host "⚠️  Orphaned Team: $($Team.DisplayName)" -ForegroundColor Red
    }
}

$OrphanedTeams | Export-Csv -Path ".\reports\Teams-Orphaned.csv" -NoTypeInformation
Write-Host "`n📊 Found $($OrphanedTeams.Count) orphaned Teams" -ForegroundColor Cyan
```

### 📄 Script 4 — Full Governance Report
```powershell
# 04-GenerateGovernanceReport.ps1
# Description: Generates a consolidated HTML governance report for all Teams

Connect-MgGraph -Scopes "Group.Read.All", "User.Read.All"

$AllTeams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All
$ReportData = @()

foreach ($Team in $AllTeams) {
    $Members = Get-MgGroupMember -GroupId $Team.Id -All
    $Owners  = Get-MgGroupOwner  -GroupId $Team.Id -All
    $Guests  = $Members | Where-Object { $_.AdditionalProperties["userType"] -eq "Guest" }

    $ReportData += [PSCustomObject]@{
        TeamName    = $Team.DisplayName
        Visibility  = $Team.Visibility
        Members     = $Members.Count
        Owners      = $Owners.Count
        Guests      = $Guests.Count
        IsOrphaned  = ($Owners.Count -eq 0)
        HasGuests   = ($Guests.Count -gt 0)
        CreatedDate = $Team.CreatedDateTime
    }
}

# Generate HTML Report
$HTML = @"
<!DOCTYPE html>
<html>
<head><title>Teams Governance Report</title>
<style>
  body { font-family: Arial; padding: 20px; }
  table { border-collapse: collapse; width: 100%; }
  th { background-color: #0078d4; color: white; padding: 10px; }
  td { border: 1px solid #ddd; padding: 8px; }
  tr:nth-child(even) { background-color: #f2f2f2; }
  .orphaned { background-color: #ffcccc !important; }
</style></head>
<body>
<h1>Microsoft Teams Governance Report</h1>
<p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')</p>
<p>Total Teams: $($ReportData.Count)</p>
<table>
<tr><th>Team Name</th><th>Visibility</th><th>Members</th><th>Owners</th><th>Guests</th><th>Orphaned?</th></tr>
"@

foreach ($Row in $ReportData) {
    $CssClass = if ($Row.IsOrphaned) { ' class="orphaned"' } else { '' }
    $HTML += "<tr$CssClass><td>$($Row.TeamName)</td><td>$($Row.Visibility)</td><td>$($Row.Members)</td><td>$($Row.Owners)</td><td>$($Row.Guests)</td><td>$($Row.IsOrphaned)</td></tr>`n"
}

$HTML += "</table></body></html>"
$HTML | Out-File -FilePath ".\reports\Teams-GovernanceReport.html" -Encoding UTF8

Write-Host "✅ HTML Report saved to .\reports\Teams-GovernanceReport.html" -ForegroundColor Green
```

---

## PROJECT 3
## 📧 Exchange Online Management Toolkit {#project-3}

### 📌 Problem It Solves
Managing mailboxes, retention policies, and bulk user provisioning in Exchange Online manually is risky and slow. This toolkit automates the most common Exchange admin tasks.

### 🛠️ Technologies Used
- Exchange Online PowerShell (EXO V3)
- Retention Policies
- Mailbox Permissions
- Mail Flow Rules

### 📂 Repository Structure
```
Exchange-Online-Toolkit/
├── README.md
├── scripts/
│   ├── 01-BulkMailboxProvisioning.ps1
│   ├── 02-RetentionPolicyAutomation.ps1
│   ├── 03-MailboxPermissionsAudit.ps1
│   └── 04-MailboxSizeReport.ps1
└── csv-templates/
    └── mailboxes-template.csv
```

### ⚙️ Prerequisites
```powershell
# Install Exchange Online module
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force

# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com
```

### 📄 Script 1 — Bulk Mailbox Provisioning
```powershell
# 01-BulkMailboxProvisioning.ps1
# Description: Provisions shared or user mailboxes in bulk from CSV
# CSV Headers: DisplayName, Alias, PrimarySmtpAddress, Type (User/Shared)

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath
)

Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com

$Mailboxes = Import-Csv -Path $CsvPath
$Results   = @()

foreach ($Mailbox in $Mailboxes) {
    try {
        if ($Mailbox.Type -eq "Shared") {
            New-Mailbox -Shared `
                -Name $Mailbox.DisplayName `
                -Alias $Mailbox.Alias `
                -PrimarySmtpAddress $Mailbox.PrimarySmtpAddress

            Write-Host "✅ Shared Mailbox Created: $($Mailbox.DisplayName)" -ForegroundColor Green
        }
        else {
            # For user mailboxes, the user must already exist in Entra ID
            Enable-Mailbox -Identity $Mailbox.PrimarySmtpAddress
            Write-Host "✅ User Mailbox Enabled: $($Mailbox.DisplayName)" -ForegroundColor Green
        }

        $Results += [PSCustomObject]@{
            DisplayName = $Mailbox.DisplayName
            Email       = $Mailbox.PrimarySmtpAddress
            Type        = $Mailbox.Type
            Status      = "Success"
        }
    }
    catch {
        Write-Host "❌ Failed: $($Mailbox.DisplayName) — $($_.Exception.Message)" -ForegroundColor Red
        $Results += [PSCustomObject]@{
            DisplayName = $Mailbox.DisplayName
            Email       = $Mailbox.PrimarySmtpAddress
            Type        = $Mailbox.Type
            Status      = "Failed: $($_.Exception.Message)"
        }
    }
}

$Results | Export-Csv -Path ".\Mailbox-Provisioning-Report.csv" -NoTypeInformation
Write-Host "`n📊 Report saved." -ForegroundColor Cyan
```

### 📄 Script 2 — Retention Policy Automation
```powershell
# 02-RetentionPolicyAutomation.ps1
# Description: Creates and assigns retention policies to mailboxes by department

Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com

# Create retention tags (if they don't exist)
$TagName = "CORP-7Year-Retention"
$Existing = Get-RetentionPolicyTag -Identity $TagName -ErrorAction SilentlyContinue

if (-not $Existing) {
    New-RetentionPolicyTag -Name $TagName `
        -Type All `
        -RetentionAction DeleteAndAllowRecovery `
        -RetentionEnabled $true `
        -AgeLimitForRetention 2555  # 7 years in days

    Write-Host "✅ Retention Tag created: $TagName" -ForegroundColor Green
}

# Create Retention Policy
$PolicyName = "CORP-Finance-RetentionPolicy"
$ExistingPolicy = Get-RetentionPolicy -Identity $PolicyName -ErrorAction SilentlyContinue

if (-not $ExistingPolicy) {
    New-RetentionPolicy -Name $PolicyName `
        -RetentionPolicyTagLinks $TagName

    Write-Host "✅ Retention Policy created: $PolicyName" -ForegroundColor Green
}

# Assign policy to Finance department mailboxes
$FinanceMailboxes = Get-Mailbox -ResultSize Unlimited -Filter {Department -eq "Finance"}

foreach ($Mailbox in $FinanceMailboxes) {
    Set-Mailbox -Identity $Mailbox.Identity -RetentionPolicy $PolicyName
    Write-Host "📋 Assigned policy to: $($Mailbox.DisplayName)" -ForegroundColor Cyan
}

Write-Host "`n✅ Retention policies applied to $($FinanceMailboxes.Count) Finance mailboxes." -ForegroundColor Green
```

### 📄 Script 3 — Mailbox Permissions Audit
```powershell
# 03-MailboxPermissionsAudit.ps1
# Description: Audits all mailbox full access and send-as permissions in the tenant

Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com

$AllMailboxes = Get-Mailbox -ResultSize Unlimited
$PermReport   = @()

foreach ($Mailbox in $AllMailboxes) {
    # Full Access Permissions
    $FullAccess = Get-MailboxPermission -Identity $Mailbox.Identity |
        Where-Object { $_.User -notlike "NT AUTHORITY*" -and $_.IsInherited -eq $false }

    # Send As Permissions
    $SendAs = Get-RecipientPermission -Identity $Mailbox.Identity |
        Where-Object { $_.Trustee -notlike "NT AUTHORITY*" }

    foreach ($Perm in $FullAccess) {
        $PermReport += [PSCustomObject]@{
            Mailbox        = $Mailbox.PrimarySmtpAddress
            PermissionType = "Full Access"
            GrantedTo      = $Perm.User
            AccessRights   = $Perm.AccessRights -join ", "
        }
    }

    foreach ($Perm in $SendAs) {
        $PermReport += [PSCustomObject]@{
            Mailbox        = $Mailbox.PrimarySmtpAddress
            PermissionType = "Send As"
            GrantedTo      = $Perm.Trustee
            AccessRights   = "SendAs"
        }
    }
}

$PermReport | Export-Csv -Path ".\Mailbox-Permissions-Audit.csv" -NoTypeInformation
Write-Host "✅ Permissions audit complete. $($PermReport.Count) permission entries found." -ForegroundColor Green
```

### 📄 Script 4 — Mailbox Size Report
```powershell
# 04-MailboxSizeReport.ps1
# Description: Generates a mailbox size report sorted by largest first
# Useful for identifying storage issues and planning migrations

Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com

$AllMailboxes = Get-Mailbox -ResultSize Unlimited
$SizeReport   = @()

foreach ($Mailbox in $AllMailboxes) {
    $Stats = Get-MailboxStatistics -Identity $Mailbox.Identity

    # Parse size string to MB
    $SizeString = $Stats.TotalItemSize.ToString()

    $SizeReport += [PSCustomObject]@{
        DisplayName    = $Mailbox.DisplayName
        Email          = $Mailbox.PrimarySmtpAddress
        MailboxType    = $Mailbox.RecipientTypeDetails
        TotalSize      = $SizeString
        ItemCount      = $Stats.ItemCount
        LastLogonTime  = $Stats.LastLogonTime
        Department     = $Mailbox.Department
    }
}

# Sort by item count descending
$SizeReport = $SizeReport | Sort-Object -Property ItemCount -Descending

$SizeReport | Export-Csv -Path ".\Mailbox-Size-Report.csv" -NoTypeInformation
Write-Host "✅ Mailbox size report exported for $($SizeReport.Count) mailboxes." -ForegroundColor Green
```

---

## PROJECT 4
## 🛡️ DLP & Compliance Auditing Tool {#project-4}

### 📌 Problem It Solves
Security teams need visibility into potential data loss prevention violations and sensitive data exposure across M365. This tool automates DLP policy auditing and generates actionable reports.

### 🛠️ Technologies Used
- Microsoft Purview (Compliance Portal)
- DLP Policy Management
- Security & Compliance PowerShell
- Audit Logs

### 📂 Repository Structure
```
DLP-Compliance-Auditing/
├── README.md
├── scripts/
│   ├── 01-AuditDLPPolicies.ps1
│   ├── 02-GetDLPViolationReport.ps1
│   ├── 03-SensitiveDataScan.ps1
│   └── 04-ComplianceScoreReport.ps1
└── reports/
    └── (generated reports appear here)
```

### ⚙️ Prerequisites
```powershell
# Install Security & Compliance module
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force

# Connect to Security & Compliance Center
Connect-IPPSSession -UserPrincipalName admin@yourdomain.com
```

### 📄 Script 1 — Audit All DLP Policies
```powershell
# 01-AuditDLPPolicies.ps1
# Description: Lists all DLP policies, their status, and associated rules

Connect-IPPSSession -UserPrincipalName admin@yourdomain.com

$DLPPolicies = Get-DlpCompliancePolicy
$PolicyReport = @()

foreach ($Policy in $DLPPolicies) {
    $Rules = Get-DlpComplianceRule -Policy $Policy.Name

    foreach ($Rule in $Rules) {
        $PolicyReport += [PSCustomObject]@{
            PolicyName       = $Policy.Name
            PolicyMode       = $Policy.Mode
            PolicyEnabled    = $Policy.Enabled
            RuleName         = $Rule.Name
            SensitiveTypes   = ($Rule.ContentContainsSensitiveInformation.Name -join "; ")
            NotifyUser       = ($Rule.NotifyUser -join "; ")
            BlockAccess      = $Rule.BlockAccess
            GenerateAlert    = $Rule.GenerateAlert
        }
    }
}

$PolicyReport | Export-Csv -Path ".\reports\DLP-Policy-Audit.csv" -NoTypeInformation
Write-Host "✅ DLP Policy Audit complete: $($DLPPolicies.Count) policies found." -ForegroundColor Green
```

### 📄 Script 2 — DLP Violation Report
```powershell
# 02-GetDLPViolationReport.ps1
# Description: Pulls DLP violation events from the audit log (last 30 days)

Connect-IPPSSession -UserPrincipalName admin@yourdomain.com

$StartDate = (Get-Date).AddDays(-30)
$EndDate   = Get-Date

Write-Host "🔍 Searching audit log for DLP violations (last 30 days)..." -ForegroundColor Cyan

$DLPEvents = Search-UnifiedAuditLog `
    -StartDate $StartDate `
    -EndDate $EndDate `
    -RecordType ComplianceDLPSharePoint, ComplianceDLPExchange `
    -ResultSize 5000

$ViolationReport = @()

foreach ($Event in $DLPEvents) {
    $AuditData = $Event.AuditData | ConvertFrom-Json

    $ViolationReport += [PSCustomObject]@{
        Timestamp       = $Event.CreationDate
        User            = $Event.UserIds
        Operation       = $Event.Operations
        Workload        = $AuditData.Workload
        PolicyName      = $AuditData.PolicyDetails.PolicyName
        RuleName        = $AuditData.PolicyDetails.Rules.RuleName
        SensitiveType   = $AuditData.SensitiveInfoDetectionIsIncluded
        ObjectName      = $AuditData.ObjectId
    }
}

$ViolationReport | Export-Csv -Path ".\reports\DLP-Violations.csv" -NoTypeInformation
Write-Host "✅ Found $($ViolationReport.Count) DLP violation events." -ForegroundColor Yellow
```

### 📄 Script 3 — Sensitive Data Location Scan
```powershell
# 03-SensitiveDataScan.ps1
# Description: Identifies locations where sensitive data types have been detected

Connect-IPPSSession -UserPrincipalName admin@yourdomain.com

# Run a content search for sensitive data types
$SearchName = "SensitiveDataScan-$(Get-Date -Format 'yyyyMMdd')"

# Common sensitive info types
$SensitiveTypes = @(
    "Credit Card Number",
    "U.S. Social Security Number (SSN)",
    "International Banking Account Number (IBAN)",
    "U.S. Individual Taxpayer Identification Number (ITIN)"
)

Write-Host "🔍 Creating content search for sensitive data types..." -ForegroundColor Cyan

New-ComplianceSearch -Name $SearchName `
    -ContentMatchQuery "SensitiveType:`"Credit Card Number`" OR SensitiveType:`"U.S. Social Security Number (SSN)`"" `
    -ExchangeLocation All `
    -SharePointLocation All `
    -OneDriveLocation All

Start-ComplianceSearch -Identity $SearchName

Write-Host "⏳ Search started. Check results in the Compliance Portal." -ForegroundColor Yellow
Write-Host "   Portal: https://compliance.microsoft.com/contentsearch" -ForegroundColor Cyan
Write-Host "   Search Name: $SearchName" -ForegroundColor Cyan

# Wait and get results
Start-Sleep -Seconds 30
$SearchResults = Get-ComplianceSearch -Identity $SearchName

Write-Host "`n📊 Search Status: $($SearchResults.Status)" -ForegroundColor Green
Write-Host "   Items Found: $($SearchResults.Items)" -ForegroundColor Green
Write-Host "   Size: $($SearchResults.Size)" -ForegroundColor Green
```

### 📄 Script 4 — Compliance Score Report
```powershell
# 04-ComplianceScoreReport.ps1
# Description: Exports a summary of current DLP policy coverage and gaps

Connect-IPPSSession -UserPrincipalName admin@yourdomain.com
Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com

$Report = @()

# Check DLP policies coverage
$AllPolicies = Get-DlpCompliancePolicy

$ExchangeCovered    = $AllPolicies | Where-Object { $_.ExchangeLocation.Count -gt 0 }
$SharePointCovered  = $AllPolicies | Where-Object { $_.SharePointLocation.Count -gt 0 }
$OneDriveCovered    = $AllPolicies | Where-Object { $_.OneDriveLocation.Count -gt 0 }
$TeamsCovered       = $AllPolicies | Where-Object { $_.TeamsLocation.Count -gt 0 }

$Report += [PSCustomObject]@{ Area = "Exchange Online"; PoliciesApplied = $ExchangeCovered.Count;   Status = if ($ExchangeCovered.Count -gt 0) {"✅ Covered"} else {"❌ Not Covered"} }
$Report += [PSCustomObject]@{ Area = "SharePoint";      PoliciesApplied = $SharePointCovered.Count; Status = if ($SharePointCovered.Count -gt 0) {"✅ Covered"} else {"❌ Not Covered"} }
$Report += [PSCustomObject]@{ Area = "OneDrive";        PoliciesApplied = $OneDriveCovered.Count;   Status = if ($OneDriveCovered.Count -gt 0) {"✅ Covered"} else {"❌ Not Covered"} }
$Report += [PSCustomObject]@{ Area = "Microsoft Teams"; PoliciesApplied = $TeamsCovered.Count;      Status = if ($TeamsCovered.Count -gt 0) {"✅ Covered"} else {"❌ Not Covered"} }

$Report | Format-Table -AutoSize
$Report | Export-Csv -Path ".\reports\DLP-Coverage-Summary.csv" -NoTypeInformation
Write-Host "✅ Compliance coverage report exported." -ForegroundColor Green
```

---

## PROJECT 5
## 📱 Intune Device Management Dashboard {#project-5}

### 📌 Problem It Solves
IT admins need a quick snapshot of device compliance, app deployment status, and security posture across all enrolled devices. This toolkit generates that visibility automatically.

### 🛠️ Technologies Used
- Microsoft Graph PowerShell SDK
- Microsoft Intune (MDM)
- Device Compliance Policies
- Security Posture Reporting

### 📂 Repository Structure
```
Intune-Device-Management/
├── README.md
├── scripts/
│   ├── 01-DeviceComplianceReport.ps1
│   ├── 02-AppDeploymentStatus.ps1
│   ├── 03-NonCompliantDevices.ps1
│   └── 04-SecurityPostureReport.ps1
└── reports/
    └── (generated reports appear here)
```

### ⚙️ Prerequisites
```powershell
# Install Microsoft Graph module
Install-Module Microsoft.Graph -Scope CurrentUser -Force

# Connect with Intune scopes
Connect-MgGraph -Scopes `
    "DeviceManagementManagedDevices.Read.All", `
    "DeviceManagementApps.Read.All", `
    "DeviceManagementConfiguration.Read.All"
```

### 📄 Script 1 — Device Compliance Report
```powershell
# 01-DeviceComplianceReport.ps1
# Description: Generates a full compliance status report for all Intune-enrolled devices

Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All"

$Devices = Get-MgDeviceManagementManagedDevice -All `
    -Property "DeviceName,ComplianceState,OperatingSystem,OsVersion,UserPrincipalName,LastSyncDateTime,ManagementAgent,Manufacturer,Model"

$Report = @()

foreach ($Device in $Devices) {
    $Report += [PSCustomObject]@{
        DeviceName       = $Device.DeviceName
        User             = $Device.UserPrincipalName
        OS               = $Device.OperatingSystem
        OSVersion        = $Device.OsVersion
        Manufacturer     = $Device.Manufacturer
        Model            = $Device.Model
        ComplianceState  = $Device.ComplianceState
        ManagementAgent  = $Device.ManagementAgent
        LastSync         = $Device.LastSyncDateTime
    }
}

# Summary stats
$Total      = $Report.Count
$Compliant  = ($Report | Where-Object { $_.ComplianceState -eq "compliant" }).Count
$NonCompliant = ($Report | Where-Object { $_.ComplianceState -eq "noncompliant" }).Count
$Unknown    = ($Report | Where-Object { $_.ComplianceState -eq "unknown" }).Count

Write-Host "`n📊 Device Compliance Summary" -ForegroundColor Cyan
Write-Host "   Total Enrolled : $Total"
Write-Host "   ✅ Compliant   : $Compliant"
Write-Host "   ❌ Non-Compliant: $NonCompliant"
Write-Host "   ⚠️  Unknown     : $Unknown"

$Report | Export-Csv -Path ".\reports\Device-Compliance-Report.csv" -NoTypeInformation
Write-Host "`n✅ Report saved to .\reports\Device-Compliance-Report.csv" -ForegroundColor Green
```

### 📄 Script 2 — App Deployment Status
```powershell
# 02-AppDeploymentStatus.ps1
# Description: Reports on Intune app deployment status across all devices

Connect-MgGraph -Scopes "DeviceManagementApps.Read.All"

# Get all managed apps
$Apps = Get-MgDeviceAppManagementMobileApp -All

$AppReport = @()

foreach ($App in $Apps) {
    # Get install summary
    $Summary = Get-MgDeviceAppManagementMobileAppInstallSummary -MobileAppId $App.Id -ErrorAction SilentlyContinue

    if ($Summary) {
        $AppReport += [PSCustomObject]@{
            AppName               = $App.DisplayName
            AppType               = $App.AdditionalProperties["@odata.type"]
            Publisher             = $App.Publisher
            InstalledDeviceCount  = $Summary.InstalledDeviceCount
            FailedDeviceCount     = $Summary.FailedDeviceCount
            PendingDeviceCount    = $Summary.PendingInstallDeviceCount
            NotInstalledCount     = $Summary.NotInstalledDeviceCount
        }
    }
}

$AppReport | Sort-Object -Property FailedDeviceCount -Descending |
    Export-Csv -Path ".\reports\App-Deployment-Status.csv" -NoTypeInformation

Write-Host "✅ App deployment report saved for $($AppReport.Count) apps." -ForegroundColor Green
```

### 📄 Script 3 — Non-Compliant Devices Alert
```powershell
# 03-NonCompliantDevices.ps1
# Description: Finds non-compliant devices and sends a summary alert via email

Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "Mail.Send"

$NonCompliant = Get-MgDeviceManagementManagedDevice -All `
    -Filter "complianceState eq 'noncompliant'" `
    -Property "DeviceName,UserPrincipalName,OperatingSystem,LastSyncDateTime,ComplianceState"

Write-Host "⚠️  Non-Compliant Devices Found: $($NonCompliant.Count)" -ForegroundColor Red

$TableRows = ""
foreach ($Device in $NonCompliant) {
    $TableRows += "<tr><td>$($Device.DeviceName)</td><td>$($Device.UserPrincipalName)</td><td>$($Device.OperatingSystem)</td><td>$($Device.LastSyncDateTime)</td></tr>`n"
}

# Build HTML email body
$EmailBody = @"
<html><body>
<h2>⚠️ Non-Compliant Devices Report</h2>
<p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')</p>
<p>Total Non-Compliant: <strong>$($NonCompliant.Count)</strong></p>
<table border='1' style='border-collapse:collapse'>
<tr style='background:#d9534f;color:white'>
  <th>Device Name</th><th>User</th><th>OS</th><th>Last Sync</th>
</tr>
$TableRows
</table>
</body></html>
"@

# Send alert email via Microsoft Graph
$Message = @{
    Subject = "⚠️ Intune Non-Compliant Devices Report - $(Get-Date -Format 'yyyy-MM-dd')"
    Body    = @{
        ContentType = "HTML"
        Content     = $EmailBody
    }
    ToRecipients = @(
        @{ EmailAddress = @{ Address = "itadmin@yourdomain.com" } }
    )
}

Send-MgUserMail -UserId "admin@yourdomain.com" -Message $Message
Write-Host "✅ Alert email sent to itadmin@yourdomain.com" -ForegroundColor Green

$NonCompliant | Export-Csv -Path ".\reports\NonCompliant-Devices.csv" -NoTypeInformation
```

### 📄 Script 4 — Security Posture Report
```powershell
# 04-SecurityPostureReport.ps1
# Description: Generates a security posture report showing encryption,
# antivirus, firewall, and OS update status across all devices

Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All"

$Devices = Get-MgDeviceManagementManagedDevice -All `
    -Property "DeviceName,UserPrincipalName,OperatingSystem,IsEncrypted,AntivirusScanRequired,IsSupervised,ComplianceState,OsVersion"

$SecurityReport = @()

foreach ($Device in $Devices) {
    $SecurityReport += [PSCustomObject]@{
        DeviceName          = $Device.DeviceName
        User                = $Device.UserPrincipalName
        OS                  = $Device.OperatingSystem
        OSVersion           = $Device.OsVersion
        IsEncrypted         = $Device.IsEncrypted
        AntivirusRequired   = $Device.AntivirusScanRequired
        IsSupervised        = $Device.IsSupervised
        ComplianceState     = $Device.ComplianceState
        EncryptionStatus    = if ($Device.IsEncrypted) {"✅ Encrypted"} else {"❌ NOT Encrypted"}
    }
}

# Summary
$NotEncrypted = ($SecurityReport | Where-Object { $_.IsEncrypted -eq $false }).Count
Write-Host "`n🔒 Security Posture Summary" -ForegroundColor Cyan
Write-Host "   Total Devices       : $($SecurityReport.Count)"
Write-Host "   ❌ Not Encrypted    : $NotEncrypted"
Write-Host "   ✅ Encrypted        : $($SecurityReport.Count - $NotEncrypted)"

$SecurityReport | Export-Csv -Path ".\reports\Security-Posture-Report.csv" -NoTypeInformation
Write-Host "`n✅ Security posture report saved." -ForegroundColor Green
```

---

## 🚀 How to Use This Portfolio on GitHub

### Step 1 — Create 5 Separate Repos
Create one GitHub repository for each project:
- `EntraID-Automation-Suite`
- `Teams-Governance-Dashboard`
- `Exchange-Online-Management-Toolkit`
- `DLP-Compliance-Auditing-Tool`
- `Intune-Device-Management-Dashboard`

### Step 2 — Add a README.md to Each Repo
Each README should include:
- What the project does
- Prerequisites and modules required
- How to run the scripts
- Screenshot or sample output (CSV/HTML)

### Step 3 — Pin All 5 Repos on Your GitHub Profile
Go to your GitHub profile → click **Customize your pins** → select all 5 repos.

### Step 4 — Create a Profile README
Create a special repo named exactly `YOUR-USERNAME/YOUR-USERNAME` with a `README.md` — this becomes your GitHub profile page. Include your M365 skills, certifications, and links to each project.

---

## 📜 Certifications to Add to Your Profile
- MS-102: Microsoft 365 Administrator
- SC-300: Identity and Access Administrator
- MS-500: Microsoft 365 Security Administrator
- MD-102: Endpoint Administrator
- AI-900: Azure AI Fundamentals *(in progress)*

---

## 📬 Contact
**LinkedIn:** https://www.linkedin.com/in/ballo-841688203
**Location:** Philippines | Open to Remote Worldwide

---
*This portfolio demonstrates real-world Microsoft 365 administration skills built over 5+ years of hands-on experience.*
