# Developed by Stian Kvia - Customer Support UA
# Verson 1.030624

# Following Commands are used
# Get-MailboxFolderPermission -Identity EMAIL:\Calendar
# Set-MailboxFolderPermission -Identity EMAIL:\Calendar -User Default -AccessRights LimitedDetails (Or some other AccessRight)
# Add-MailboxFolderPermission -Identity EMAIL:\Calendar -User SomeUser -AccessRights 

# Beginning of Script
Write-host "Beginning of PowerShell Script" -ForegroundColor Green
Write-Host ""

# Check if $MeetingroomFrom is empty
Write-host "Set copy settings" -NoNewline
Write-host " From " -ForegroundColor Green -NoNewline
Write-host "Meetingroom"

if ([string]::IsNullOrEmpty($MeetingroomFrom)) {
    $meetingRoomMessage = "Meeting room is not set, please enter a meetingroom email address"
} else {
    $meetingRoomMessage = "Current meeting room: $MeetingroomFrom Press Enter to keep or enter a new email"
}

# Display current value and prompt user to keep or enter a new one
$MeetingroomChoice = Read-Host -Prompt "$meetingRoomMessage"

# If user chooses to enter a new value
if ($MeetingroomChoice -ne "") {
    $MeetingroomFrom = $MeetingroomChoice
    Write-Host "Meeting room '$MeetingroomFrom' is set. Continuing with the script..." -ForegroundColor Green
} else {
    # If user chooses to keep the current value or presses Enter
        if ([string]::IsNullOrEmpty($MeetingroomFrom)) {
        Write-Host "Meeting room not provided. Exiting script."
        return
    }
    # Add the specific part of the script you want to execute when $MeetingroomFrom is set
    Write-Host "Meeting room '$MeetingroomFrom' is set. Continuing with the script..." -ForegroundColor Green
}


Write-host ""
Write-host "Set copy settings" -NoNewline
Write-host " To " -ForegroundColor Green -NoNewline
Write-host "Meetingroom"

if ([string]::IsNullOrEmpty($MeetingroomTo)) {
    $meetingRoomMessage = "Meeting room is not set, please enter a meetingroom email address"
} else {
    $meetingRoomMessage = "Current meeting room: $MeetingroomTo Press Enter to keep or enter a new email"
}

# Display current value and prompt user to keep or enter a new one
$MeetingroomChoice = Read-Host -Prompt "$meetingRoomMessage"

# If user chooses to enter a new value
if ($MeetingroomChoice -ne "") {
    $MeetingroomTo = $MeetingroomChoice
    Write-Host "Meeting room '$MeetingroomTo' is set. Continuing with the script..." -ForegroundColor Green
} else {
    # If user chooses to keep the current value or presses Enter
        if ([string]::IsNullOrEmpty($MeetingroomTo)) {
        Write-Host "Meeting room not provided. Exiting script."
        return
    }
    # Add the specific part of the script you want to execute when $MeetingroomTo is set
    Write-Host "Meeting room '$MeetingroomTo' is set. Continuing with the script..." -ForegroundColor Green
}

    Write-host ""

    # MeetingroomFrom - Get-MailboxFolderPermissions from Meetingroom
    Write-host "Getting Mailbox Folder Permissions" -ForegroundColor Yellow
    $mailboxFolderPermission = Get-MailboxFolderPermission -Identity ${MeetingroomFrom}:\Calendar

    # MeetingroomFrom - Mailbox Folder Permissions - Initialize a variable to store the concatenated Set-MailboxFolderPermission commands
    $allMailboxFolderPermissionCommands = ""

    # MeetingroomFrom - Mailbox Folder Permissions - Construct the Set-MailboxFolderPermission commands
    Write-host "Configuring Permissions" -ForegroundColor Yellow
    foreach ($permission in $mailboxFolderPermission) {
        if ($permission.User -like "Anonymous" -or $permission.User -like "Default") {
            $mailboxFolderPermissionCommand = "Set-MailboxFolderPermission -Identity $($MeetingroomTo):\Calendar -User $($permission.User) -AccessRights $($permission.AccessRights)"
    } else {
        # Resolve the user's email address
        $userEmail = (Get-Mailbox -Identity $permission.User).PrimarySmtpAddress
        $mailboxFolderPermissionCommand = "Add-MailboxFolderPermission -Identity $($MeetingroomTo):\Calendar -User $($userEmail) -AccessRights $($permission.AccessRights)"
        }
    # Concatenate the commands with line breaks
        $allMailboxFolderPermissionCommands += "$mailboxFolderPermissionCommand`r`n"
    }

    # MeetingroomFrom - Execute Commands to Configure new meetingroom
$ErrorActionPreference = 'Stop'

try {
    Invoke-Expression $allMailboxFolderPermissionCommands
    Write-host "Configuration should be copied from $MeetingroomFrom to $MeetingroomTo" -ForegroundColor Green
} catch {
    if ($_.Exception.Message -like "*The specified mailbox Identity*doesn't exist*") {
        Write-Host ""
        Write-Host "The specified mailbox does not exist." -ForegroundColor Red
        Write-Host "$_" -ForegroundColor Yellow
    } elseif ($_.Exception.Message -like "*An existing permission entry was found*") {
        Write-Host ""
        Write-Host "An existing permission entry was found" -ForegroundColor Red
        Write-Host "$_" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Here is access so you can manually make sure it is correct" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Meetingroom FROM" -ForegroundColor Green
        Get-MailboxFolderPermission -Identity "${MeetingroomFrom}:\Calendar"
        Write-Host ""
        Write-Host "Meetingroom TO" -ForegroundColor Green
        Get-MailboxFolderPermission -Identity "${MeetingroomTo}:\Calendar"
    } else {
        Write-Host ""
        Write-Host "An unexpected error occurred:" -ForegroundColor Red
        Write-Host "$_" -ForegroundColor Yellow
    }
}