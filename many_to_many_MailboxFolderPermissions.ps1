# Connect to 365 exchange.
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking


function Safe-MailFolder-Perm {
    param([string[]]$give = 'None',
    [string[]]$access_to = 'None',
    [string]$with_permission = 'None',
    [string[]]$folders = 'None')

    # Fix AccessRights param.
    if($with_permission -eq 'read')
        {$with_permission = 'Reviewer'}
    ElseIf ($with_permission -eq 'write')
        {$with_permission = 'Editor'}
    ElseIf ($with_permission -eq 'all')
        {$with_permission = 'PublishingEditor'}
    ElseIf ($with_permission -eq 'CalLimitedDetails' -Or $with_permission -eq 'LimitedDetails')
        {$with_permission = 'LimitedDetails'}

    if($with_permission -eq 'LimitedDetails' -and $folders -eq 'calendar-only')
        {Write-Host "ERROR: Unable to assign 'LimitedDetails' permission to non-calendar mail folders!"; exit}

    # Fix folder perms.
    if($folders -eq 'inbox-only')
        {$folders = ':\',':\inbox'}
    ElseIf($folders -eq 'calendar-only')
        {$folders = ':\',':\calendar'}
    ElseIf($folders -eq 'both')
        {$folders = ':\',':\Inbox' ,':\Calendar'}


    # Check if multiple users provided.
    $give = $give | ForEach-Object {"$_@bmtlogin.com"}

    $access_to = $access_to | ForEach-Object {"$_@bmtlogin.com"}

    #$access_to | Select PrimarySMTPAddress -ExpandProperty PrimarySMTPAddress

    foreach($give_address in $give){
        foreach($access_address in $access_to){
            Write-Host $access_address
            foreach($folder in $folders){
                # These if statements use Add- as Set- does not work if there is no existing folder permission for that user.
                    Set-MailboxFolderPermission $access_address${folder} -User ${give_address} -AccessRights ${with_permission}

                    if($?) # Explaination of '$?':  https://stackoverflow.com/questions/10634115/what-is-in-powershell
                    {
                        Write-Host "SUCCESS: Setting mailbox access for:" + ${give_address} + " on behalf of " + ${access_address}+${folder}
                    }
                    else
                    {
                        Write-Host "INFO: Failed to use 'Set-' command, attempting permission 'Add-'..."
                        Add-MailboxFolderPermission $access_address${folder} -User ${give_address} -AccessRights ${with_permission}
                        if($?)
                        {
                            Write-Host "SUCCESS: Set root mailbox access for:" + ${give_address} + " on behalf of " + ${access_address}+${folder}
                        }
                        else
                        {
                            Write-Host "WARNING: Failed setting root mailbox permission for: " + ${give_address} + ' to access ' + ${access_address}${folder}
                        }
                    }
                }
            }
        }
    }


Write-Host "STARTING ..."

Safe-MailFolder-Perm -give 'xxxxxx' -access_to 'xxxxx' -with_permission write -folders both

# Disconnect from 365, otherwise we could be limited by max concurrent sessions.
Remove-PSSession $Session
