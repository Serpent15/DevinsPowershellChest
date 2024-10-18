# Import the Exchange Online Management module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
Connect-ExchangeOnline 

# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Prepare results array
$results = @()

foreach ($mbx in $mailboxes) {

    # Try to get mailbox statistics
    try {
        # Use UserPrincipalName to ensure uniqueness
        $stats = Get-MailboxStatistics -Identity $mbx.UserPrincipalName

        # Get TotalItemSize in bytes
        if ($stats.TotalItemSize -and $stats.TotalItemSize.Value -ne $null) {
            $totalItemSizeString = $stats.TotalItemSize.Value.ToString()
            if ($totalItemSizeString -match '\((\d+(?:,\d+)*)\sbytes\)') {
                $totalItemSizeInBytes = [long]$matches[1].Replace(',', '')
            } else {
                $totalItemSizeInBytes = 0
            }
        } else {
            $totalItemSizeInBytes = 0
        }

        # Get ProhibitSendReceiveQuota in bytes
        if ($mbx.ProhibitSendReceiveQuota -and $mbx.ProhibitSendReceiveQuota.IsUnlimited) {
            $quotaInBytes = 50GB
        } elseif ($mbx.ProhibitSendReceiveQuota -and $mbx.ProhibitSendReceiveQuota.Value -ne $null) {
            $quotaValue = $mbx.ProhibitSendReceiveQuota.Value
            $quotaString = $quotaValue.ToString()
            if ($quotaString -match '(\d+(?:\.\d+)?)\s*(MB|GB|TB)') {
                $sizeValue = [double]$matches[1]
                $sizeUnit = $matches[2]
                switch ($sizeUnit) {
                    'MB' { $quotaInBytes = $sizeValue * 1MB }
                    'GB' { $quotaInBytes = $sizeValue * 1GB }
                    'TB' { $quotaInBytes = $sizeValue * 1TB }
                }
            } else {
                $quotaInBytes = 50GB # Default value if parsing fails
            }
        } else {
            $quotaInBytes = 50GB # Default value if ProhibitSendReceiveQuota is null
        }

        # Calculate percent used
        if ($quotaInBytes -gt 0) {
            $percentUsed = [math]::Round(($totalItemSizeInBytes / $quotaInBytes) * 100, 2)
        } else {
            $percentUsed = 0
        }

        # Determine if the mailbox is nearing capacity
        $nearingCapacity = if ($percentUsed -ge 80) { "Yes" } else { "No" }

        # Add to results array
        $results += [PSCustomObject]@{
            DisplayName       = $mbx.DisplayName
            UserPrincipalName = $mbx.UserPrincipalName
            TotalItemSizeGB   = [math]::Round($totalItemSizeInBytes / 1GB, 2)
            QuotaGB           = [math]::Round($quotaInBytes / 1GB, 2)
            PercentUsed       = $percentUsed
            NearingCapacity   = $nearingCapacity
        }

    } catch {
        # Handle the error and continue
        Write-Host "Error processing mailbox '$($mbx.DisplayName)': $_" -ForegroundColor Red
        continue
    }
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false

# Output the results to Out-GridView
$results | Sort-Object -Property PercentUsed -Descending | Out-GridView -Title "M365 Mailbox Sizes"
