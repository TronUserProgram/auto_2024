# Stian Kvia
$Updated = "27.10.2024"

# Beginning of Script
Write-host "Beginning of PowerShell Script" -ForegroundColor Green
Write-Host ""

########################### Check Module and Connected Customer

if ($Connect_Customer -eq "Enabled") {
} else {
    Write-host "Please Run Connect Customer PowerShell" -ForegroundColor Red
    return
}

Connected

#############################################################################################################

# Confirm correct customer
$confirmation = Read-Host "Are you sure you are connected to Exchange Online of the correct customer? (Y/N)"
if ($confirmation -eq "Y" -or $confirmation -eq "y") {

# Gather Information
$SharedMailboxName = Read-Host "Enter the Name of the Shared Mailbox"
$SharedMailbox = Read-Host "Enter the Shared Mailbox Email Address"
Write-host ""

# Extract the part before the "@" symbol
$SharedMailboxAlias = $SharedMailbox -split "@" | Select-Object -First 1

# Create Shared Mailbox and run Configuration
New-Mailbox -Shared -Name "$SharedMailboxName" -PrimarySmtpAddress "$SharedMailbox" -Alias "$SharedMailboxAlias"
Write-host ""
Write-host "Pending 10 Seconds before config running"
Timeout /T 10
Set-Mailbox -Identity "$SharedMailbox" -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true
Write-host ""

# Define the giveaccess function
function GiveAccess {
    param (
        [string]$email,
        [string]$access = "FullAccess"
    )

    # Define the commands based on the value of $access
    switch ($access) {
        "FullAccess" {
            $command = "Add-MailboxPermission -Identity $SharedMailbox -User $email -AccessRights FullAccess; Add-RecipientPermission -Identity $SharedMailbox -AccessRights SendAs -Trustee $email -Confirm:`$false"
            Write-Host "Access granted to $email on $SharedMailbox."
        }
        "SendAs" {
            $command = "Add-RecipientPermission $SharedMailbox -AccessRights SendAs -Trustee $email -Confirm:`$false"
            Write-Host "Granted SendAs Access to $email on $SharedMailbox."
        }
        Default {
            Write-Host "Invalid access type."
            return
        }
    }

    # Execute the command using Invoke-Expression
    Invoke-Expression $command
}

Write-host ""
Write-host ""
Write-host ""
Write-host ""
Write-host "You can add users manually to this Shared Mailbox with the command GiveAccess user@domain.com" -ForegroundColor Yellow

} else {
Write-host ""
Write-host "If you have used Connect Customer PowerShell you can connect with:"
Write-host "Connect Exchange `"Customer`""
    Return
}
