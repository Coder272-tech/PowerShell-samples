# Parameters
param (
    [string]$TenantId = "your-tenant-id",
    [string]$ClientId = "your-client-id",
    [string]$ClientSecret = "your-client-secret",
    [string]$SiteId = "your-sharepoint-site-id",
    [string]$LibraryId = "your-document-library-id",
    [string]$ArchiveFolder = "Archive",
    [int]$DaysOld = 30,
    [string]$AdminEmail = "admin@example.com"
)

# Authenticate with Microsoft Graph
Write-Host "Authenticating with Microsoft Graph..."
$AuthBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $ClientId
    Client_Secret = $ClientSecret
}
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $AuthBody
$AccessToken = $TokenResponse.access_token

# Get files from SharePoint document library
Write-Host "Fetching files from SharePoint Online..."
$Files = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$LibraryId/root/children" `
    -Headers @{ Authorization = "Bearer $AccessToken" } -Method GET

# Process files
$CurrentDate = Get-Date
$OldFiles = @()
foreach ($File in $Files.value) {
    if ($File.lastModifiedDateTime -lt ($CurrentDate.AddDays(-$DaysOld))) {
        $OldFiles += $File
        Write-Host "Moving file: $($File.name)"
        
        # Move file to Archive folder
        $MoveUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$LibraryId/items/$($File.id)/move"
        $MoveBody = @{
            parentReference = @{
                path = "/drives/$LibraryId/root:/$ArchiveFolder"
            }
        } | ConvertTo-Json -Depth 2
        Invoke-RestMethod -Uri $MoveUri -Headers @{ Authorization = "Bearer $AccessToken" } -Method POST -Body $MoveBody
    }
}

# Prepare email summary
Write-Host "Sending email summary..."
$EmailBody = @"
The following files have been moved to the '$ArchiveFolder' folder in SharePoint Online:
$($OldFiles.name -join "`n")
"@

# Send email
Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me/sendMail" `
    -Headers @{ Authorization = "Bearer $AccessToken" } `
    -Method POST -Body @{
        message = @{
            subject = "SharePoint Archive Summary"
            body    = @{
                contentType = "Text"
                content     = $EmailBody
            }
            toRecipients = @(@{ emailAddress = @{ address = $AdminEmail } })
        }
    } | ConvertTo-Json -Depth 3

Write-Host "Process completed successfully."
