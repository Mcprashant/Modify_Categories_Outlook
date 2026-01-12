# ===========================
# INTERACTIVE CONFIGURATION
# ===========================
Write-Host "Please provide the required configuration:" -ForegroundColor Yellow

$TenantId = Read-Host "Enter Tenant ID"
$ClientId = Read-Host "Enter Client ID"
$ClientSecretSecure = Read-Host "Enter Client Secret" -AsSecureString
$ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecretSecure))
$GroupId = Read-Host "Enter Group ID"

# Interactive categories to delete
Write-Host "`nEnter categories to delete (one per line, Enter blank line to finish):" -ForegroundColor Cyan
$CategoriesToDelete = @()
do {
    $cat = Read-Host "Category name"
    if ($cat -and $cat.Trim() -ne "") {
        $CategoriesToDelete += $cat.Trim()
    }
} while ($cat -ne "")

if ($CategoriesToDelete.Count -eq 0) {
    Write-Host "No categories specified. Exiting." -ForegroundColor Red
    exit
}

Write-Host "Categories to delete: $($CategoriesToDelete -join ', ')"

# ===========================
# AUTHENTICATION
# ===========================
Write-Host "Authenticating with Microsoft Graph..."
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -Method POST -ContentType "application/x-www-form-urlencoded" `
    -Body @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

$AccessToken = $TokenResponse.access_token
$Headers = @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" }

# ===========================
# GET GROUP MEMBERS
# ===========================
Write-Host "Fetching members of group $GroupId..."
$Members = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/members?`$select=userPrincipalName" `
    -Headers $Headers -Method GET

# Handle pagination
while ($Members.'@odata.nextLink') {
    $NextPage = Invoke-RestMethod -Uri $Members.'@odata.nextLink' -Headers $Headers
    $Members.value += $NextPage.value
    $Members.'@odata.nextLink' = $NextPage.'@odata.nextLink'
}

Write-Host "Found $($Members.value.Count) members."

# ===========================
# DELETE CATEGORIES
# ===========================
foreach ($member in $Members.value) {
    if ($member.userPrincipalName) {
        Write-Host "`nProcessing user: $($member.userPrincipalName)"

        try {
            # Get existing categories
            $ExistingCategories = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($member.userPrincipalName)/outlook/masterCategories" `
                -Headers $Headers -Method GET

            foreach ($cat in $ExistingCategories.value | Where-Object { $_.displayName -in $CategoriesToDelete }) {
                try {
                    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($member.userPrincipalName)/outlook/masterCategories/$($cat.id)" `
                        -Method DELETE -Headers $Headers
                    Write-Host "🗑 Deleted category '$($cat.displayName)'"
                }
                catch {
                    Write-Host "❌ Failed to delete category '$($cat.displayName)': $($_.Exception.Message)"
                }
            }
        }
        catch {
            Write-Host "❌ Failed to retrieve categories for $($member.userPrincipalName): $($_.Exception.Message)"
        }
    }
}

Write-Host "`n✅ Completed deletion for all group members."

# Clear sensitive variables from memory
$ClientSecret = $null
$AccessToken = $null
[GC]::Collect()
