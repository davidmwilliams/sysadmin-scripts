#Requires -Version 5.1
<#

  _   _                          _____                                           _ 
 | \ | |                        |  __ \                                         | |
 |  \| |  ___ __      __ ______ | |__) |__ _  ___  ___ __      __ ___   _ __  __| |
 | . ` | / _ \\ \ /\ / /|______||  ___// _` |/ __|/ __|\ \ /\ / // _ \ | '__|/ _` |
 | |\  ||  __/ \ V  V /         | |   | (_| |\__ \\__ \ \ V  V /| (_) || |  | (_| |
 |_| \_| \___|  \_/\_/          |_|    \__,_||___/|___/  \_/\_/  \___/ |_|   \__,_|
                                                                                                                                                                   

#>
# -------- HELP --------
<#
.Credit
    ALL credit goes to MAGC. 
    Code commented and documented by JVM
.Synopsis
    This script will generate a new secure password string or credentialobject     
.PARAMETER AsString
    Specify if return object should be plaintext string
.PARAMETER Length
    Specifies the length of the 
.PARAMETER ForbiddenChars
    Allows user to make specific chars forbiden
.PARAMETER MinLowerCaseChars
    Set minimum amount of required lower case chars
.PARAMETER MinUpperCaseChars
    Set minimum amount of upper case chars
.PARAMETER MinDigits
    Set minimum amount of digits
.PARAMETER MinSpecialChars
    Set minimum amount of special chars required
#>
function New-Password
{
    [CmdletBinding(PositionalBinding = $false)]
    [Alias("np")]
    [OutputType([securestring],[string])]
    
    #--------------------------------------------| PARAMETERS |--------------------------------------------#
    Param
    (
        [Parameter()]
        [switch]
        $AsString,
        
        [Parameter()]
        [ValidateRange(8,[int]::MaxValue)]
        [Int]
        $Length=40,
        
        [Parameter()]
        [Alias("DisallowedChars")]
        [ArgumentCompleter(
            {
                param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
                switch ($wordToComplete -replace "`"|'")
                {
                    {"Lowercase" -like "$_*"}
                    {
                        [System.Management.Automation.CompletionResult]::new(
                            "abcdefghijklmnopqrstuvwxyz".ToCharArray().ForEach({"'$_'"}) -join ',',
                            'Lowercase',
                            [System.Management.Automation.CompletionResultType]::ParameterValue,
                            'Lowercase'
                        )
                    }
                    {"Uppercase" -like "$_*"}
                    {
                        [System.Management.Automation.CompletionResult]::new(
                            "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray().ForEach({"'$_'"}) -join ',',
                            'Uppercase',
                            [System.Management.Automation.CompletionResultType]::ParameterValue,
                            'Uppercase'
                        )
                    }
                    {"Digits" -like "$_*"}
                    {
                        [System.Management.Automation.CompletionResult]::new(
                            "1234567890".ToCharArray().ForEach({"'$_'"}) -join ',',
                            'Digits',
                            [System.Management.Automation.CompletionResultType]::ParameterValue,
                            'Digits'
                        )
                    }
                    {"Special" -like "$_*"}
                    {
                        [System.Management.Automation.CompletionResult]::new(
                            '/*!\"$%()=?{[]}+#-.,<_:;>~|@'.ToCharArray().ForEach({"'$_'"}) -join ',',
                            'Special',
                            [System.Management.Automation.CompletionResultType]::ParameterValue,
                            'Special'
                        )
                    }
                    {"Ambiguous" -like "$_*"}
                    {
                        [System.Management.Automation.CompletionResult]::new(
                            "IlOo0".ToCharArray().ForEach({"'$_'"}) -join ',',
                            'Ambiguous',
                            [System.Management.Automation.CompletionResultType]::ParameterValue,
                            'Ambiguous'
                        )
                    }
                }
            }
        )]
        [char[]]
        $ForbiddenChars,
        
        [Parameter()]
        [ValidateRange(0,[int]::MaxValue)]
        [Int]
        $MinLowercaseChars=2,
        
        [Parameter()]
        [ValidateRange(0,[int]::MaxValue)]
        [Int]
        $MinUppercaseChars=2,
        
        [Parameter()]
        [ValidateRange(0,[int]::MaxValue)]
        [Int]
        $MinDigits=2,
        
        [Parameter()]
        [ValidateRange(0,[int]::MaxValue)]
        [Int]
        $MinSpecialChars=2
    )
    #---------------------------------------------| CHECK INPUT |--------------------------------------------#
    begin
    {
        # Start out by building $AllAllowedChars variable. This is all subvariables concatinated, where no forbidden chars are included
        [char[]]$AllAllowedChars = @(
            ([char[]]$AllowedLowercase = "abcdefghijklmnopqrstuvwxyz".ToCharArray().Where({$_ -cnotin $ForbiddenChars}))
            ([char[]]$AllowedUppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray().Where({$_ -cnotin $ForbiddenChars}))
            ([char[]]$AllowedDigits    = "1234567890".ToCharArray().Where({$_ -notin $ForbiddenChars}))
            ([char[]]$AllowedSpecial   = '/*!\"$%()=?{[]}+#-.,<_:;>~|@'.ToCharArray().Where({$_ -notin $ForbiddenChars}))
        )
        # FillerCharCount refers to the amount of characters not dictated by the required minimum of each type       
        [int]$FillerCharCount = $Length - ($MinLowercaseChars + $MinUppercaseChars + $MinDigits + $MinSpecialChars)
        
        # For all if statements below, throw erorr if minimum requirements not met.
        if ($FillerCharCount -lt 0)
        {
            throw "The specified length is less than the sum of the minimum character counts."
        }
        if ($AllowedLowercase.Count -lt 1 -and $MinLowercaseChars -gt 0)
        {
            throw "There are not enough allowed lowercase chars for the specified minimum lowercase count."
        }
        if ($AllowedUppercase.Count -lt 1 -and $MinUppercaseChars -gt 0)
        {
            throw "There are not enough allowed uppercase chars for the specified minimum uppercase count."
        }
        if ($AllowedDigits.Count -lt 1 -and $MinDigits -gt 0)
        {
            throw "There are not enough allowed digits for the specified minimum digit count."
        }
        if ($AllowedSpecial.Count -lt 1 -and $MinSpecialChars -gt 0)
        {
            throw "There are not enough allowed special chars for the specified minimum special count."
        }
        # Function to generate random chars for array. Takes the chararray to populate and an amount as input
        function GetRandomChars ([char[]]$CharArray, [int]$Amount)
        {
            # Check if input is valid
            if ($CharArray.Count -gt 0 -and $Amount -gt 0)
            {
                # Fills array with random chars from input array
                for ($i = 0; $i -lt $Amount; $i++)
                {
                    $CharArray[(Get-Random -Maximum $CharArray.Count)]
                }
            }
        }
    }
    #------------------------------------------| BUILD PASSWORD |------------------------------------------#
    process
        {
            try
            {
                if ($AsString)
                {
                    # User wants output as plain text string
                    $StringBuilder = [System.Text.StringBuilder]::new($Length)
                }
                else
                {
                    # User want output as secure string
                    $SecureString = [securestring]::new()
                }
                
                # Get all random chars in fixed position with GetRandomChars function
                # Randomize their order with Get-Random, 
                # for each char either append to secure string or plain text string depending on user choice
                @(
                    GetRandomChars -CharArray $AllowedLowercase -Amount $MinLowercaseChars
                    GetRandomChars -CharArray $AllowedUppercase -Amount $MinUppercaseChars
                    GetRandomChars -CharArray $AllowedDigits    -Amount $MinDigits
                    GetRandomChars -CharArray $AllowedSpecial   -Amount $MinSpecialChars
                    GetRandomChars -CharArray $AllAllowedChars  -Amount $FillerCharCount
                ) | 
                Get-Random -Count $Length | ForEach-Object -Process {
                    if ($AsString)
                    {
                        $null = $StringBuilder.Append($_)
                    }
                    else
                    {
                        $SecureString.AppendChar($_)
                    }
                }
                # Entire pass wword has been built
                if ($AsString)
                {
                    # Return plaintext string if user asks for that
                    $StringBuilder.ToString()
                }
                else
                {
                    # return secure string if user did not ask for cleartext
                    $SecureString
                }
            }
            catch
            {
                Write-Error $_
            }
        }
    }
