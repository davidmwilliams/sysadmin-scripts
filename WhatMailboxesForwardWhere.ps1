# 1. Install and Import the Exchange Online Module if missing
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}
Import-Module ExchangeOnlineManagement

# 2. Connect to Exchange Online
Connect-ExchangeOnline

# 3. Fetch all mailboxes and filter for active forwarding
Write-Host "Fetching mailboxes and checking forwarding status..." -ForegroundColor Cyan

# Using Get-EXOMailbox for optimized performance in large tenants
$ForwardingMailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward | 
    Where-Object { $_.ForwardingAddress -ne $null -or $_.ForwardingSmtpAddress -ne $null }


# 4. Process and format the results
$Report = foreach ($Mailbox in $ForwardingMailboxes) {
    [PSCustomObject]@{
        "DisplayName"               = $Mailbox.DisplayName
        "UserPrincipalName"         = $Mailbox.UserPrincipalName
        "ForwardingAddress"         = $Mailbox.ForwardingAddress
        "ForwardingSmtpAddress"     = $Mailbox.ForwardingSmtpAddress
        "DeliverToMailboxAndForward"= $Mailbox.DeliverToMailboxAndForward
    }
}

# 5. Export results to CSV
if ($Report) {
    $OutputPath = "C:\Temp\M365-MailboxForwardingReport.csv"
    # Ensure local directory exists
    if (-not (Test-Path "C:\Temp")) { New-Item -ItemType Directory -Path "C:\Temp" | Out-Null }
    
    $Report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    Write-Host "Success! Report saved to: $OutputPath" -ForegroundColor Green
    $Report | Out-GridView -Title "Mailboxes with Forwarding Enabled"
} else {
    Write-Host "No mailboxes found with mailbox-level forwarding enabled." -ForegroundColor Yellow
}

# 6. Disconnect Session
Disconnect-ExchangeOnline -Confirm:$false
