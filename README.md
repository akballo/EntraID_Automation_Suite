## 🔐 Entra ID Automation Suite

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

