$properties = @(
    @{n='TimeStamp';e={$_.TimeCreated}},
    @{n='Who did it';e={$_.Properties[6].Value}},
    @{n='Reason';e={$_.Properties[0].Value}},
    @{n='Action';e={$_.Properties[4].Value}}
)
Get-WinEvent -FilterHashTable @{LogName='System'; ID=1074} | 
Select $properties | Sort-Object "$_.TimeCreated"
