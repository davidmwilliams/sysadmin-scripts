$Computer = Read-Host 'Servername'
$Port = Read-Host 'Port'
Test-NetConnection -ComputerName $Computer -Port $Port |
Select @{Name="Servername";Expression={$_.Computername}}, @{Name="IP";Expression={$_.Remoteaddress}}, @{Name="Is $port open?";Expression={$_.TcpTestSucceeded}}
