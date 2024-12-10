# Automated Active Directory User Report Generator
# Requires: Active Directory module

# Parameters
param (
    [string]$OutputFile = "AD_User_Report.csv",
    [switch]$SendEmail,
    [string]$EmailRecipient = "admin@example.com",
    [string]$SMTPServer = "smtp.example.com"
)

# Check if Active Directory module is available
if (-not (Get-Module -ListAvailable -Name "ActiveDirectory")) {
    Write-Error "Active Directory module not found. Please install it before running this script."
    exit
}

# Import Active Directory module
Import-Module ActiveDirectory

# Fetch all Active Directory users
try {
    $users = Get-ADUser -Filter * -Property DisplayName, Enabled, LastLogonDate, MemberOf
} catch {
    Write-Error "Failed to fetch Active Directory users. $_"
    exit
}

# Process and format data
$userData = foreach ($user in $users) {
    [PSCustomObject]@{
        DisplayName   = $user.DisplayName
        Enabled       = $user.Enabled
        LastLogonDate = $user.LastLogonDate
        Groups        = ($user.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }) -join ";"
    }
}

# Export to CSV
try {
    $userData | Export-Csv -Path $OutputFile -NoTypeInformation -Force
    Write-Host "User report generated successfully at $OutputFile"
} catch {
    Write-Error "Failed to export user data to CSV. $_"
    exit
}

# Optional: Email the report
if ($SendEmail) {
    try {
        Send-MailMessage -From "noreply@example.com" -To $EmailRecipient -Subject "AD User Report" `
            -Body "Please find the attached Active Directory user report." `
            -Attachments $OutputFile -SmtpServer $SMTPServer
        Write-Host "Email sent successfully to $EmailRecipient"
    } catch {
        Write-Error "Failed to send email. $_"
    }
}
