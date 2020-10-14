<#
.SYNOPSIS
  Gets a random strong password
.DESCRIPTION
  Gets a random strong password
.EXAMPLE
  PS> get-randomStrongPassword -Length 10
  u+,ob-_"N[a
#>
function get-randomStrongPassword {
    [CmdletBinding()]
    [OutputType([string])]
    param(
      [Parameter(Mandatory=$false)][int]$Length = 21
    )
    
    ([char[]]([char]33..[char]95) + ([char[]]([char]97..[char]126)) + 0..9 | Sort-Object {Get-Random})[0..$Length] -join '' 
  }


get-randomStrongPassword