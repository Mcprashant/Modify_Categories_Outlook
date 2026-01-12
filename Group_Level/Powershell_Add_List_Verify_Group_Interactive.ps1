# ===========================
# INTERACTIVE CONFIGURATION
# ===========================
Write-Host "Please provide the required configuration:" -ForegroundColor Yellow

$TenantId = Read-Host "Enter Tenant ID"
$ClientId = Read-Host "Enter Client ID"
$ClientSecretSecure = Read-Host "Enter Client Secret" -AsSecureString
$ClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecretSecure))
$GroupId = Read-Host "Enter Group ID"
$DeleteAfterVerify = Read-Host "Delete after verify? (y/n)" -eq 'y'

# Categories to manage
$Categories = @(
    @{ displayName = "Vergadering intern"; color = "preset3" },
    @{ displayName = "Vergadering extern"; color = "preset7" },
    @{ displayName = "Focuswerk";          color = "preset10" },
    @{ displayName = "Afwezigheid";        color = "preset5" }
)

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
# PROCESS EACH USER
# ===========================
foreach ($member in $Members.value) {
    if ($member.userPrincipalName) {
        Write-Host "`nProcessing user: $($member.userPrincipalName)"

        # STEP 1: CREATE CATEGORIES
        foreach ($cat in $Categories) {
            $Body = $cat | ConvertTo-Json
            try {
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($member.userPrincipalName)/outlook/masterCategories" `
                    -Method POST -Headers $Headers -Body $Body
                Write-Host "✅ Added category '$($cat.displayName)'"
            }
            catch {
                Write-Host "❌ Failed to add category '$($cat.displayName)': $($_.Exception.Message)"
            }
        }

        # STEP 2: VERIFY CATEGORIES
        try {
            $ExistingCategories = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($member.userPrincipalName)/outlook/masterCategories" `
                -Headers $Headers -Method GET
            Write-Host "Current categories for $($member.userPrincipalName):"
            $ExistingCategories.value | Format-Table displayName, color
        }
        catch {
            Write-Host "❌ Failed to list categories: $($_.Exception.Message)"
        }

        # STEP 3: OPTIONAL DELETE
        if ($DeleteAfterVerify) {
            Write-Host "Deleting categories for $($member.userPrincipalName)..."
            foreach ($cat in $Categories) {
                try {
                    $catToDelete = $ExistingCategories.value | Where-Object { $_.displayName -eq $cat.displayName }
                    if ($catToDelete) {
                        Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$($member.userPrincipalName)/outlook/masterCategories/$($catToDelete.id)" `
                            -Method DELETE -Headers $Headers
                        Write-Host "🗑️ Deleted category '$($cat.displayName)'"
                    }
                }
                catch {
                    Write-Host "⚠️ Failed to delete category '$($cat.displayName)': $($_.Exception.Message)"
                }
            }
        }
    }
}

Write-Host "`n✅ Completed processing all group members."

# Clear sensitive variables from memory
$ClientSecret = $null
$AccessToken = $null
[GC]::Collect()
