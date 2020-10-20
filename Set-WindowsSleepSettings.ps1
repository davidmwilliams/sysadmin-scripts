param (
    # Seconds to wait before sleeping the computer when running on battery.
    #
    # Defaults to 3600 seconds, or one hour.
    [int] $DcTimeout = 3600,

    # Seconds to wait before sleeping the computer while plugged in.
    #
    # Defaults to 0 which indicated that the computer should never sleep.
    [int] $AcTimeOut = 0
)

# Get the current PowerPlan (e.g. Balanced or High Performance)
$ActiveScheme = Invoke-Expression 'powercfg /getactivescheme'

# Turn the output into something usable (RegEx magic)
$RegEx = '(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}'
$Guid = [regex]::Match($ActiveScheme,$RegEx).Value

# ID for the category and setting we want to change
# https://docs.microsoft.com/en-us/windows-hardware/customize/power-settings/configure-power-settings
$SleepSettings = '238c9fa8-0aad-41ed-83f4-97be242c8f20'
$IdleTimeout = '29f6c1db-86da-48c5-9fdb-f2b67b1f44da'

# Change the "DC" setting, so while on battery. 3600 = Sleep after one hour idle
Invoke-Expression "powercfg /setdcvalueindex $Guid $SleepSettings $IdleTimeout $DcTimeout"

# Change the "AC" setting, so while on wall power. 0 = Never sleep
Invoke-Expression "powercfg /setacvalueindex $Guid $SleepSettings $IdleTimeout $AcTimeOut"

# Save the changes
Invoke-Expression "powercfg /s $Guid"
