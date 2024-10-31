######################################################
#                                                    #
#  Developed by Stian Kvia | Customer Support - UA   #
#                                                    #
######################################################

### SETTING PARAMETERS
Param(

    [Parameter(HelpMessage = "Trigger after xx Days|Set how many days since the user was disabled before triggering automation", Mandatory = $true)]
    [String]$days = "30",                                           # Default Value

    [Parameter(HelpMessage = "Remove UserData|", Mandatory = $true)]
    [ValidateSet('Enable', 'Disable')]
    [String]$RemoveUserData = "Enable",                             # Default is Enable

    [Parameter(HelpMessage = "UserData Path|")]
    [string]$SourcePath_UserData = "[UserDataRoot]",                # Default Value

    [Parameter(HelpMessage = "Remove Citrix Profile|", Mandatory = $true)]
    [ValidateSet('Enable', 'Disable')]
    [String]$RemoveCitrix = "Disable",                              # Default is Disable

    [Parameter(HelpMessage = "Citrix Profile Path||RemoveCitrix==='Enable'")]
    [string]$SourcePath_Citrix = "\\[DNSDomainName]\root\Citrix",   # Default Value

    [Parameter(HelpMessage = "Remove UPD Disk|", Mandatory = $true)]
    [ValidateSet('Enable', 'Disable')]
    [string]$RemoveUPD = "Disable",                                 # Default is Disable

    [Parameter(HelpMessage = "UPD Disk Path||RemoveUPD==='Enable'")]
    [string]$SourcePath_UPD = "[CitrixCloudUPD]",                   # Default Value

    [Parameter(HelpMessage = "Remove GeoCloud UPD Disk|", Mandatory = $true)]
    [ValidateSet('Enable', 'Disable')]
    [string]$RemoveGeoCloudUPD = "Disable",                         # Default is Disable

    [Parameter(HelpMessage = "GeoCloud UPD Disk Path||RemoveGeoCloudUPD==='Enable'")]
    [string]$SourcePath_GeoCloudUPD = "[CitrixCloudUPD]",           # Default Value
    #[string]$SourcePath_GeoCloudUPD = "[GeoCloudUPD]",             # Default Value

    [Parameter(HelpMessage = "Move OU|Move user to Deleted CCP", Mandatory = $true)]
    [ValidateSet('Enable', 'Disable')]
    [string]$MoveOU = "Enable",                                     # Default is Enable

    [Parameter(HelpMessage = "Delete User|Can only be used if Move OU is Disabled", Mandatory = $true)]
    [ValidateSet('Enable', 'Disable')]
    [string]$DeleteUser = "Disable",                                # Default is Disable

    [Parameter(HelpMessage = "OU Path - Disabled Users|OU to search for disabled users")]
    [String]$AD_Disabled_PathOU = "[ADDisabledUsersOU]",            # Default Value

    [Parameter(HelpMessage = "OU Path - Deleted CCP|")]
    [String]$AD_DeletedCCP_PathOU = "OU=Deleted CCP,[ADBaseDN]",    # Default Value

    [Parameter(HelpMessage = "Domain Name|Should be default value.", Mandatory = $false)]
    [string]$DNSDomainName = "[DNSDomainName]"                           # Default Value

    

)

$Source = "Offboarding_Log"
$EventLog = "Offboarding_UA_Test"

$Message = "Script Start"       
LogEventInfo -EventID 1000 -Message $Message
Write-Output $Message

#LogEventInfo
#LogEventWarning
#LogEventError

# Set if logging should be Enable or Disabled, and What location
$logging = "Enable"
$logFile = "C:\scripts\Offboarding.log"

#### PENDING TO SET THE CLOUD LOGGING


##################################### AUTOMATION #####################################

# Check Logging Settings
if ($Logging -eq "Enable") {
    # Test Log File Path
    if (-not (Test-Path $logFile)) {
        # Create CSV File and Headers
        New-Item -Path $logFile -ItemType File -Force
        Add-Content -Path $logFile -Value "Date,Name,DaysSinceDisabled,DisabledDate,UserData,UPD,GeoCloudUPD,CitrixData,MoveOU,DeleteUser"
    }
} # End Logging Check

# Function to get SID from samAccountName
function Get-SIDFromSamAccountName($samAccountName) {
    try {
        # Retrieve the user object from Active Directory
        $user = Get-ADUser -Identity $samAccountName -Properties SID
        return $user.SID.Value
    } catch {
        $Message = "Error retrieving SID for $samAccountName"
        LogEventError -EventID 2000 -Message $Message
        Write-Output $Message
        return $null
    }
}

# Function User Data - Set Folder Permissions
# Can be used as | Set-FolderPermissions -Path UserData Path or Any other path
# This has to run BEFORE Set-FolderPermissions Function
function Set-FolderPermissions {
    param (
    [string]$Path = $FolderPath
    )

    # Assign the input parameter to the internal variable
    $FolderPath = $Path

    try {
        # Take ownership of the folder
        takeown /F "$FolderPath" /R /D Y
    
        # Get the current ACL
        $acl = Get-Acl -Path $FolderPath

        # Create the new access rule
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $domainUser,
            [System.Security.AccessControl.FileSystemRights]::FullControl,
            [System.Security.AccessControl.InheritanceFlags]::ContainerInherit,
            [System.Security.AccessControl.PropagationFlags]::None,
            [System.Security.AccessControl.AccessControlType]::Allow
        )

        # Add the access rule to the ACL
        $acl.SetAccessRule($accessRule)
    
        # Apply the updated ACL
        Set-Acl -Path $FolderPath -AclObject $acl -ErrorAction Stop

        Write-Output "Ownership and permissions set successfully for $FolderPath"
    } catch {
        Write-Output "Error trying to take ownership of folder: $($_.Exception.Message)"
    }
}

# Function User Data - Set File Permissions
# Can be used as | Set-FilePermissions -Path UserData Path or Any other path
# This has to run AFTER Set-FolderPermissions Function
function Set-FilePermissions {
    param (
    [string]$Path = $FolderPath
    )
    
    # Assign the input parameter to the internal variable
    $FolderPath = $Path

    try {
    $Set_FilePermissions = "icacls `"$FolderPath`" /grant `"$env:USERNAME`:(F)`"/T"
    Invoke-Expression $Set_FilePermissions
    } catch {
        Write-Output "Error trying to set file permissions: $($_.Exception.Message)"
    }
}

# Function Secondary Solution Remove Data
# Can be used as | Set-FilePermissions -Path UserData Path or Any other path
# This has to run AFTER Set-FolderPermissions Function
function Remove-Data {
    try {
        # Remove directories under $FolderPath
        Get-ChildItem -Path $FolderPath -Recurse -Directory -Force | ForEach-Object {
            try {
                Remove-Item -LiteralPath "\\?\$($_.FullName)" -Recurse -Force -ErrorAction Continue
            } catch {
                Write-Output "Failed to remove directory $($_.FullName): $($_.Exception.Message)"
            }
        }

        # Optionally remove the root folder itself if needed
        #Remove-Item -LiteralPath "\\?\$FolderPath" -Recurse -Force -ErrorAction Stop

        # Update status and message if successful
        $userDataStatus = "Removed"
        $Message = "Removing UserData Success - Secondary Solution"
        Write-Output $Message

    } catch [System.UnauthorizedAccessException] {
        Write-Output "Access denied when trying to remove $FolderPath. Check permissions."
        $Message = "Removing UserData Failed - Secondary Solution"
        Write-Output $Message
    } catch {
        Write-Output "An error occurred while trying to remove $FolderPath : $($_.Exception.Message)"
        $Message = "Removing UserData Failed - Secondary Solution"
        Write-Output $Message
    }
}



# Define the date threshold (get value from days)
$dateThreshold = (Get-Date).AddDays(-$days)

# Get all disabled users in the Disabled OU
$disabledUsers = Get-ADUser -Filter { Enabled -eq $false } -SearchBase $AD_Disabled_PathOU -Properties whenChanged, HomeDirectory, samAccountName
$Message = "Getting Disabled Users $disabledUsers"       
LogEventInfo -EventID 1010 -Message $Message
Write-Output $Message

foreach ($user in $disabledUsers) {
    $samAccountName = $user.samAccountName
    $userDN = $user.DistinguishedName

    $Message = "Checking User: $samAccountName, $userDN"       
    LogEventInfo -EventID 1020 -Message $Message
    # Write-Output $Message

    # Set empty variables
    #$date = $null
    $name = $null
    $daysSinceDisabled = $null
    $disableDate = $null
    $userDataFolderStatus = $null
    $UPDFileStatus = $null
    $userCitrixFolderStatus = $null
    $moveUserStatus = $null
    $deleteUserStatus = $null

    # Get the SID for the specified samAccountName
    $sid = Get-SIDFromSamAccountName -samAccountName $samAccountName
    
    # Get the metadata for the user object using repadmin
    $repadminOutput = & repadmin /showobjmeta "$DNSDomainName" $userDN
    
    # Parse the repadmin output to find the userAccountControl attribute change
    $lines = $repadminOutput -split "`n"
    $userAccountControlLine = $lines | Where-Object { $_ -match "userAccountControl" }

# UPDATE 31.10.24 _ Replace due to object duplication issues if there is multiple DC with different dates
    if ($userAccountControlLine) {
        # Parse the date from the lines and sort by the date
        $sortedUserAccountControlLine = $userAccountControlLine | ForEach-Object {
            $parts = $_ -split "\s+"
            $dateString = $parts | Where-Object { $_ -match '^\d{4}-\d{2}-\d{2}$' }
            $timeString = $parts | Where-Object { $_ -match '^\d{2}:\d{2}:\d{2}$' }
            [pscustomobject]@{
                Line = $_
                DateTime = [datetime]::ParseExact("$dateString $timeString", 'yyyy-MM-dd HH:mm:ss', $null)
            }
        } | Sort-Object DateTime -Descending
        
        # Get the most recent entry
        $mostRecentUserAccountControlLine = $sortedUserAccountControlLine | Select-Object -First 1
        
        # Extract the date and time parts
        $dateString = $mostRecentUserAccountControlLine.Line -split "\s+" | Where-Object { $_ -match '^\d{4}-\d{2}-\d{2}$' }
        $timeString = $mostRecentUserAccountControlLine.Line -split "\s+" | Where-Object { $_ -match '^\d{2}:\d{2}:\d{2}$' }
        
        # Combine date and time into a single string
        $dateTimeString = "$dateString $timeString"

        # Convert the string to a DateTime object
        $accountControlDate = [datetime]::ParseExact($dateTimeString, "yyyy-MM-dd HH:mm:ss", $null)
        
        # Check the condition and set the message
        if ($accountControlDate -gt $dateThreshold) {
            $Message = "SKIP USER - $samAccountName was Disabled less than $days days ago - Date: $dateString"
        } else {
            $Message = "PROCESS USER - $samAccountName was Disabled more than $days days ago - Date: $dateString"
        }
   
        LogEventInfo -EventID 1030 -Message $Message
        Write-Output $Message

        # Check if the disable date is older than * days
        if ($accountControlDate -lt $dateThreshold) {
            Write-Output "Automation Triggered for $samAccountName"
            # Construct the user's data path based on samAccountName

####################################### REMOVE USER DATA
            if ($RemoveUserData -eq "Enable") {
                $Message = "Remove UserData Triggered"
                $userDataPath = $samAccountName
                # Define the source and destination paths based on samAccountName
                $userDataFolder = Join-Path -Path $SourcePath_UserData -ChildPath $userDataPath

                # Check if the user's folder exists in the source path
                if (Test-Path -Path $userDataFolder) {
                    # Delete the UserData\UserFolder
                    try {
                        Remove-Item -Path $userDataFolder -Recurse -Force
                        $userDataStatus = "Removed"
                        $Message = "Removing UserData Success"                       
                    } catch {
                        $userDataStatus = "Failed"
                        $Message = "Removing UserData Failed"
                        $FolderPath = $userDataFolder
                        Set-FolderPermissions
                        Set-FilePermissions
                        Remove-Data
                    }
                } else {
                    $userDataStatus = "No Folder"
                    $Message = "No UserData Folder"
                }
                LogEventInfo -EventID 1040 -Message $Message
                Write-Output $Message
                $userDataFolderStatus = $userDataStatus
            }
            
##########################################################################################################################

####################################### REMOVE FS Logix Citrix Profile
            if ($RemoveCitrix -eq "Enable") {
                $Message = "Remove Citrix Triggered"
                $userDataPath = $samAccountName
                # Define the source and destination paths based on samAccountName
                $userCitrixFolder = Join-Path -Path $SourcePath_Citrix -ChildPath $userDataPath

                # Check if the user's folder exists in the source path
                if (Test-Path -Path $userCitrixFolder) {
                    # Delete the UserData\UserFolder
                    Write-Output "Citrix Folder Triggered"
                    try {
                        Remove-Item -Path $userCitrixFolder -Recurse -Force
                        $userDataStatus = "Removed"
                        $Message = "Removing Citrix Folder Success"
                    } catch {
                        $userDataStatus = "Failed"
                        $Message = "Removing Citrix Folder Failed"
                    }
                } else {
                    Write-Output "Folder for user $samAccountName does not exist in $SourcePath_Citrix"
                    $userDataStatus = "No Folder"
                    $Message = "No Citrix Folder"
                }
                LogEventInfo -EventID 1050 -Message $Message
                Write-Output $Message
                $userCitrixFolderStatus = $userDataStatus
            }
##########################################################################################################################

####################################### REMOVE UPD Disk
## Adding some more outputs here as there is some issue setting the message from past output
            if ($RemoveUPD -eq "Enable") {
                $Message = "Remove UPD Disk Triggered"
                if ($SourcePath_UPD) {
                    Write-Output "Sourcepath UPD: $SourcePath_UPD"
                    if ($sid) {
                        Write-Output "SID: $sid"
                        # Define the pattern to search for VHDX files that include the SID
                        $filePattern = "UVHD-*$sid*.vhdx"
                        Write-Output "Filepattern: $filePattern"
            
                        # Get list of VHDX files matching the pattern
                        $vhds = Get-ChildItem -Path $SourcePath_UPD -Filter $filePattern
                        Write-Output "VHDS: $vhds"

                        # Initialize a flag to check if any folders were found
                        $UPDFound = $false

                        foreach ($vhd in $vhds) {
                            $UPDFound = $true
                            if (Test-Path $vhd.FullName) {
                                try {
                                    $Message = "Remove UPD Failed ($_.Exception.Message)"
                                    Remove-Item -Path $vhd.FullName -Force
                                    $userDataStatus = "Removed"
                                    $Message = "Remove UPD Success"
                                } catch {
                                    $userDataStatus = "Failed"
                                    $Message = "Remove UPD Failed - $($vhd.FullName) - $($_.Exception.Message)"
                                }
                            } else {
                                $userDataStatus = "No UPD"
                                $Message = "No UPD Found - $($vhd.FullName)"
                            }
                        }

                        # If no folders were found, set the appropriate message
                        if (-not $UPDFound) {
                            $userDataStatus = "Not Found"
                            $Message = "No UPD Found for SID: $sid"
                        }
                    } else {
                        $userDataStatus = "SID Not Found"
                        $Message = "SID Not Found"
                    }
                }
                LogEventInfo -EventID 1060 -Message $Message
                Write-Output $Message
                $UPDFileStatus = $userDataStatus
            }
##########################################################################################################################

####################################### REMOVE GeoCloud UPD Disk
            if ($RemoveGeoCloudUPD -eq "Enable") {
                 if ($SourcePath_GeoCloudUPD) {
                    if ($sid) {
                        # Define the pattern to search for folders that include the SID
                        $folderPattern = "*$sid*"

                        # Get list of folders matching the pattern
                        $folders = Get-ChildItem -Path $SourcePath_GeoCloudUPD -Directory | Where-Object { $_.Name -like $folderPattern }

                        # Initialize a flag to check if any folders were found
                        $GeoCloudUPDFound = $false

                        foreach ($folder in $folders) {
                            $GeoCloudUPDFound = $true
                            if (Test-Path $folder.FullName) {
                                try {
                                    Remove-Item -Path $folder.FullName -Recurse -Force
                                    $userDataStatus = "Removed"
                                    $Message = "Remove GeoCloud UPD Success"
                                } catch {
                                    $userDataStatus = "Failed"
                                    $Message = "Remove GeoCloud UPD Failed - $($folder.FullName) - $_.Exception.Message"
                                }
                            } else {
                                $userDataStatus = "No UPD"
                                $Message = "No GeoCloud UPD Found - $($folder.FullName)"
                            }
                        }
        
                        # If no folders were found, set the appropriate message
                        if (-not $GeoCloudUPDFound) {
                            $userDataStatus = "Not Found"
                            $Message = "No GeoCloud UPD Folders Found for SID: $sid"
                        }
                    } else {
                        $userDataStatus = "SID Not Found"
                        $Message = "SID Not Found"
                    }
                }
                Write-Output $Message
                $GeoCloudUPDFileStatus = $userDataStatus
            }
##########################################################################################################################

####################################### Move the user to the Deleted CCP OU
            if ($MoveOU -eq "Enable") {
                try {
                # Move-ADObject -Identity $user -TargetPath $AD_DeletedCCP_PathOU # Did not work
                Move-ADObject -Identity $user.DistinguishedName -TargetPath $AD_DeletedCCP_PathOU
                $moveUserStatus = "Moved"
                $Message = "Move OU Success"
                } catch {
                $moveUserStatus = "Failed"
                $Message = "Move OU Failed"
                }
                LogEventInfo -EventID 1080 -Message $Message
                Write-Output $Message
            }
##########################################################################################################################

####################################### Delete the user
            if ($DeleteUser -eq "Enable") {
                try {
                # Delete-ADObject -Identity $user -TargetPath $AD_Disabled_PathOU # Did not work
                Remove-ADUser -Identity $user.DistinguishedName -Confirm:$false
                $deleteUserStatus = "Removed"
                $Message = "Delete User Success"
                } catch {
                Write-Output "Issue Removing User" -ForegroundColor Red
                $deleteUserStatus = "Failed"
                $Message = "Delete User Failed"
                }
                LogEventInfo -EventID 1090 -Message $Message
                Write-Output $Message
            }
##########################################################################################################################

####################################### Logging
            if ($Logging -eq "Enable") {
                $date = Get-Date
                $name = $user.Name
                $daysSinceDisabled = (New-TimeSpan -Start $accountControlDate -End $date).Days
                Add-Content -Path $logFile -Value "$date,$name,$daysSinceDisabled,$disableDate,$userDataFolderStatus,$UPDFileStatus,$GeoCloudUPDFileStatus,$userCitrixFolderStatus,$moveUserStatus,$deleteUserStatus"
                $Message = "Logging Done"
                LogEventInfo -EventID 1100 -Message $Message
                Write-Output $Message
            }
##########################################################################################################################
        }
    }
}

$Message = "Offboarding Script Done"
LogEventInfo -EventID 3000 -Message $Message
Write-Output $Message

