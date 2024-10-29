# Developed by Stian Kvia - Customer Support UA
# Verson 1.030624

# Following Commands are used
# Get-MailboxFolderPermission -Identity EMAIL:\Calendar
# Set-MailboxFolderPermission -Identity EMAIL:\Calendar -User Default -AccessRights LimitedDetails (Or some other AccessRight)
# Add-MailboxFolderPermission -Identity EMAIL:\Calendar -User SomeUser -AccessRights 

# Beginning of Script
Write-host "Beginning of PowerShell Script" -ForegroundColor Green
Write-Host ""

# Check if $SharedMailboxFrom is empty
Write-host "Set copy settings" -NoNewline
Write-host " From " -ForegroundColor Green -NoNewline
Write-host "Shared Mailbox"

if ([string]::IsNullOrEmpty($SharedMailboxFrom)) {
    $SharedMailboxMessage = "Shared Mailbox is not set, please enter a shared mailbox email address"
} else {
    $SharedMailboxMessage = "Current Shared Mailbox: $SharedMailboxFrom Press Enter to keep or enter a new email"
}

# Display current value and prompt user to keep or enter a new one
$SharedMailboxChoice = Read-Host -Prompt "$SharedMailboxMessage"

# If user chooses to enter a new value
if ($SharedMailboxChoice -ne "") {
    $SharedMailboxFrom = $SharedMailboxChoice
    Write-Host "Meeting room '$SharedMailboxFrom' is set. Continuing with the script..." -ForegroundColor Green
} else {
    # If user chooses to keep the current value or presses Enter
        if ([string]::IsNullOrEmpty($SharedMailboxFrom)) {
        Write-Host "Shared Mailbox not provided. Exiting script."
        return
    }
    # Add the specific part of the script you want to execute when $SharedMailboxFrom is set
    Write-Host "Shared Mailbox '$SharedMailboxFrom' is set. Continuing with the script..." -ForegroundColor Green
}


Write-host ""
Write-host "Set copy settings" -NoNewline
Write-host " To " -ForegroundColor Green -NoNewline
Write-host "Shared Mailbox"

if ([string]::IsNullOrEmpty($SharedMailboxTo)) {
    $SharedMailboxMessage = "Shared Mailbox is not set, please enter a shared mailbox email address"
} else {
    $SharedMailboxMessage = "Current Shared Mailbox: $SharedMailboxFrom Press Enter to keep or enter a new email"
}

# Display current value and prompt user to keep or enter a new one
$SharedMailboxChoice = Read-Host -Prompt "$SharedMailboxMessage"

# If user chooses to enter a new value
if ($SharedMailboxChoice -ne "") {
    $SharedMailboxTo = $SharedMailboxChoice
    Write-Host "Shared Mailbox '$SharedMailboxTo' is set. Continuing with the script..." -ForegroundColor Green
} else {
    # If user chooses to keep the current value or presses Enter
        if ([string]::IsNullOrEmpty($SharedMailboxTo)) {
        Write-Host "Shared Mailbox not provided. Exiting script."
        return
    }
    # Add the specific part of the script you want to execute when $SharedMailboxTo is set
    Write-Host "Shared Mailbox '$SharedMailboxTo' is set. Continuing with the script..." -ForegroundColor Green
}

    Write-host ""

    # SharedMailboxFrom - Get-MailboxFolderPermissions from Shared Mailbox
    Write-host "Getting Mailbox Folder Permissions" -ForegroundColor Yellow
    $SharedMailboxPermission = Get-MailboxPermission -Identity $SharedMailboxFrom

    # SharedMailboxFrom - Mailbox Folder Permissions - Initialize a variable to store the concatenated Set-MailboxFolderPermission commands
    $AllSharedMailboxPermissionCommands = ""



    # SharedMailboxFrom - Mailbox Folder Permissions - Construct the Add-MailboxPermission commands
Write-host "Configuring Permissions" -ForegroundColor Yellow

# Initialize the variable to store all commands
$AllSharedMailboxPermissionCommands = @()
$AllSharedRecipientPermissionCommands = @()

# Filter out "NT AUTHORITY\SELF" users
$FilteredPermissions = $SharedMailboxPermission | Where-Object { $_.User -notlike "NT AUTHORITY\SELF" }

foreach ($permission in $FilteredPermissions) {
    $SharedMailboxPermissionCommand = "Add-MailboxPermission -Identity $SharedMailboxTo -User $($permission.User) -AccessRights $($permission.AccessRights) -ErrorAction SilentlyContinue"
    $SharedRecipientPermissionCommand = "Add-RecipientPermission -Identity $SharedMailboxTo -AccessRights SendAs -Trustee $($permission.User) -ErrorAction SilentlyContinue"
    # Add the command to the array
    $AllSharedMailboxPermissionCommands += $SharedMailboxPermissionCommand
    $AllSharedRecipientPermissionCommands += $SharedRecipientPermissionCommand
}

# Convert the array of commands into a single string
$AllSharedMailboxPermissionCommandsString = $AllSharedMailboxPermissionCommands -join "`r`n"
$AllSharedRecipientPermissionCommandsString = $AllSharedRecipientPermissionCommands -join "`r`n"

# Output the string of commands
Invoke-Expression $AllSharedMailboxPermissionCommandsString
Invoke-Expression $AllSharedRecipientPermissionCommandsString
