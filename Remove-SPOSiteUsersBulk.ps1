[CmdletBinding(
    SupportsShouldProcess,
    ConfirmImpact = 'High'
)]
param (
    [Parameter(Mandatory,
        Position = 0
    )]
    [System.Object[]]
    $InputObject,

    [Parameter()]
    [string]
    $OutputDirectory,

    [Parameter()]
    [Switch]
    $Live,

    [Parameter()]
    [string[]]
    $ExcludeUser
)

begin {

    [Console]::ResetColor()

    #Region Functions
    Function LogEnd {
        $txnLog = ""
        Do {
            try {
                Stop-Transcript | Out-Null
            }
            catch [System.InvalidOperationException] {
                $txnLog = "stopped"
            }
        } While ($txnLog -ne "stopped")
    }

    Function LogStart {
        param (
            [Parameter(Mandatory = $true)]
            [string]$LogPath
        )
        LogEnd
        Start-Transcript $logPath -Force | Out-Null
    }

    Function Say {
        param(
            [Parameter(
                Mandatory,
                Position = 0
            )]
            [String]
            $What,

            [Parameter(Position = 1)]
            [ValidateSet(
                "Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", "DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", "Red", "Magenta", "Yellow", "White"
            )]
            [String]
            $Color = 'Cyan'
        )

        $Host.UI.RawUI.ForegroundColor = $Color

        "$(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') : $What" | Out-Default

        [Console]::ResetColor()
    }
    #EndRegion Functions

    ## Define console message color scheme
    $color_error = 'Red'
    $color_information = 'Cyan'
    $color_warning = 'DarkYellow'
    $color_ok = 'Green'

    $now = Get-Date
    $nowString = $now.ToString("yyyy-MM-dd_hh-mm-ss_tt")

    ## Start transaction logging

    $txnLogFilename = "SPO_User_Remove_Log_$($nowString)_$($env:USERNAME).log"
    if ($OutputDirectory) {
        $txnLogFilename = "$($OutputDirectory)\$($txnLogFilename)"
    }
    LogStart -logPath $txnLogFilename
    if (!($PSBoundParameters.ContainsKey("WhatIf"))) {
        Say "Log filename is [$(Resolve-Path $txnLogFilename)]" $color_information
    }

    # Test if SPO Shell is connected
    try {
        $null = Get-SPOTenant -ErrorAction Stop -OutVariable spoTenant
        Say "[OK] SPO shell is connected." $color_ok
    }
    catch {
        Say "The Get-SpoTenant command failed. Please connect to your SharePoint Online organization first using the Connect-SPOService cmdlet." $color_error
        ## Stop transaction logging
        LogEnd
        ## Exit the script
        Continue
    }

    ## Create output file
    $resultFilename = "SPO_User_Remove_Result_$($nowString)_$($env:USERNAME).csv"
    if ($OutputDirectory) {
        $resultFilename = "$($OutputDirectory)\$($resultFilename)"
    }

    try {
        $null = New-Item -ItemType File -Name $resultFilename -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        if (!($PSBoundParameters.ContainsKey("WhatIf"))) {
            Say "Output filename is [$(Resolve-Path $resultFilename)]" $color_information
        }
        # Say "Output filename is [$(Resolve-Path $resultFilename)]" $color_information
    }
    catch {
        ## Show error and exist the script if the result file cannot be created.
        Say $_.Exception.Message $color_error
        ## Stop transaction logging
        LogEnd
        Continue
    }

    # Define the required input fields in the $InputObject object
    Say "Inspecting input object..." $color_information
    $requiredFields = @(
        'Url', 'LoginName'
    )

    # Inspect the input fields

    ## Get the input object properties (NoteProperty)
    $inputProperties = ($InputObject | Get-Member -MemberType NoteProperty)
    $requiredFields | ForEach-Object {
        ## Check if the required field is present in the input object
        if ($inputProperties.Name -notcontains $_) {
            ## If the required input object field is missing
            Say "The required [$_] column is missing." $color_error
            ## Exist the script
            continue
        }
    }
    Say "[OK] Input object list is valid." $color_ok

    ## Group the users into sites.
    Say "Aggregating users by site..." $color_information
    $inputObjectGroupedByUrl = @($InputObject | Group-Object -Property Url | Sort-Object -Property Url)

    ## Count unique URLs
    $urlCountUnique = ($inputObjectGroupedByUrl).Count

    ## Count unique users
    $userCountUnique = ($InputObject | Select-Object -Unique -Property LoginName).Count

    Say "There are $($userCountUnique) unique users to be removed from $($urlCountUnique) sites." $color_information

    ## Initialize site counter
    $siteIndex = 0

}
process {
    foreach ($site in $inputObjectGroupedByUrl) {
        $siteIndex++

        $userCountinSite = $site.Group.LoginName.Count
        Say "Site [$($siteIndex) of $($urlCountUnique)] : $($site.Name)" $color_information
        Say "    Getting the list of users on the site..." $color_information


        ## Test if the site exists and can be retreived.
        try {
            $null = Get-SPOSite -Identity $site.Name -ErrorAction Stop
            $userCollection = (Get-SPOUser -Site $site.Name -Limit All).LoginName
        }
        catch {
            Say "    $($_.Exception.Message)" $color_error

            [PSCustomObject]([ordered]@{
                    Site      = $($site.Name)
                    LoginName = "N/A"
                    Status    = "[FAILED]"
                    Details   = $_.Exception.Message
                }) | Export-Csv -Append -NoClobber -NoTypeInformation -Path $resultFilename -Force

            ## Skip to the next site
            Continue
        }

        ## Initialize user counter
        $userIndex = 0

        ## Process each user removal
        foreach ($user in ($site.Group.LoginName | Sort-Object)) {
            $userIndex++
            Say "    Removing user [$($userIndex) of $($userCountinSite)] : $($user)" $color_information

            ## If the user is in the exclusion list, skip it.
            if ($user -in $ExcludeUser) {
                Say "        Result : [SKIPPED] - The user is in the exclusion list." $color_warning
                [PSCustomObject]([ordered]@{
                        Site      = $($site.Name)
                        LoginName = $user
                        Status    = "[SKIPPED]"
                        Details   = "The user is in the exclusion list."
                    }) | Export-Csv -Append -NoClobber -NoTypeInformation -Path $resultFilename -Force
                Continue
            }

            ## If the user does not exist in the site collection, skip it.
            if ($user -notin $userCollection) {
                Say "        Result : [SKIPPED] - The user does not exist in the site." $color_warning
                [PSCustomObject]([ordered]@{
                        Site      = $($site.Name)
                        LoginName = $user
                        Status    = "[SKIPPED]"
                        Details   = "The user does not exist in the site."
                    }) | Export-Csv -Append -NoClobber -NoTypeInformation -Path $resultFilename -Force
                Continue
            }

            try {
                if (!$Live) {

                    Say "        Result : [TEST MODE] - No change." $color_information

                    [PSCustomObject]([ordered]@{
                            Site      = $($site.Name)
                            LoginName = $user
                            Status    = "[TEST MODE]"
                            Details   = "No change."
                        }) | Export-Csv -Append -NoClobber -NoTypeInformation -Path $resultFilename -Force
                }
                else {
                    if ($PSCmdlet.ShouldProcess($user, "Remove user from $($site.Name)")) {
                        ## Remove the user if not in test mode.
                        Remove-SPOUser -Site $site.Name -LoginName $user -ErrorAction Stop

                        Say "        Result : [OK] - User removed from the site." $color_ok

                        [PSCustomObject]([ordered]@{
                                Site      = $($site.Name)
                                LoginName = $user
                                Status    = "[OK]"
                                Details   = "User removed from the site."
                            }) | Export-Csv -Append -NoClobber -NoTypeInformation -Path $resultFilename -Force
                    }
                    else {
                        Say "        Result : [SKIPPED] - The removal action was not confirmed." $color_ok

                        [PSCustomObject]([ordered]@{
                                Site      = $($site.Name)
                                LoginName = $user
                                Status    = "[SKIPPED]"
                                Details   = "The removal action was not confirmed."
                            }) | Export-Csv -Append -NoClobber -NoTypeInformation -Path $resultFilename -Force
                    }
                }
            }
            catch {
                Say "        Result : [FAILED] - $($_.Exception.Message)" $color_error
                [PSCustomObject]([ordered]@{
                        Site      = $($site.Name)
                        LoginName = $user
                        Status    = "[FAILED]"
                        Details   = $_.Exception.Message
                    }) | Export-Csv -Append -NoClobber -NoTypeInformation -Path $resultFilename -Force
            }
        }
    }
}

end {
    ## Stop transaction logging
    LogEnd
}
