# Import the Exchange Online PowerShell module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online (you will be prompted for credentials)
Connect-ExchangeOnline

# Get all user mailboxes (licensed users)
$mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited

# Extract the primary domains from the mailboxes, excluding onmicrosoft.com
$domains = $mailboxes | ForEach-Object {
    $email = $_.PrimarySmtpAddress.ToString()
    $domain = $email.Split('@')[1].ToLower()
    if ($domain -notlike '*.onmicrosoft.com') {
        $domain
    }
}

# Get the most common domain
$tenantDomain = $domains | Group-Object | Sort-Object -Property Count -Descending | Select-Object -First 1 -ExpandProperty Name

# If no domains are found (unlikely), default to 'Tenant'
if (-not $tenantDomain) {
    $tenantDomain = 'Tenant'
}

# Sanitize the tenant domain to remove invalid file name characters
$invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
$invalidCharsPattern = '[{0}]' -f ([Regex]::Escape(($invalidChars -join '')))
$sanitizedTenantName = $tenantDomain -replace $invalidCharsPattern, '_'

# Initialize an array to hold the output data
$results = @()

foreach ($mailbox in $mailboxes) {
    # Get the primary SMTP address
    $primarySmtp = $mailbox.PrimarySmtpAddress

    # Get aliases (email addresses starting with "smtp:", which are not the primary address)
    $aliases = $mailbox.EmailAddresses |
        Where-Object { $_ -like "smtp:*" -and $_ -notlike "SMTP:$primarySmtp" } |
        ForEach-Object { $_.ToString().Substring(5) }  # Remove the "smtp:" prefix

    # If no aliases, set it to 'None'
    if ($aliases.Count -eq 0) {
        $aliases = @('None')
    }

    # Create a custom object for each alias
    foreach ($alias in $aliases) {
        $results += [PSCustomObject]@{
            DisplayName        = $mailbox.DisplayName
            UserPrincipalName  = $mailbox.UserPrincipalName
            PrimarySMTPAddress = $primarySmtp
            Alias              = $alias
        }
    }
}

# Export the results to a CSV file in the current directory with the tenant name
$scriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$csvFileName = "MailboxAliases_$sanitizedTenantName.csv"
$csvPath = Join-Path -Path $scriptDirectory -ChildPath $csvFileName

$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "CSV file exported to: $csvPath"

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false
