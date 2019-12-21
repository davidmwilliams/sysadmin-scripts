$ServerName = Read-host 'Enter servername'
New-ADGroup -Name "$ServerName Remote" -Description "Remote Desktop permissions to $ServerName" -GroupCategory Security -GroupScope Global -Path "OU=Groups,DC=company,DC=com"
New-ADGroup -Name "$ServerName Admin" -Description "Local administrator permissions to $ServerName" -GroupCategory Security -GroupScope Global -Path "OU=Groups,DC=Company,DC=com"
Timeout 5
Invoke-Command $ServerName -ScriptBlock {Net Localgroup "Remote Desktop Users" "domain.com\$Using:ServerName Remote" /add}
Invoke-Command $ServerName -ScriptBlock {Net Localgroup Administrators "domain.com\$Using:ServerName Admin" /add}
