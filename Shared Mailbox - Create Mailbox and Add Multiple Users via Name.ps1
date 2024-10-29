# Developed by Stian Kvia - Customer Support UA
$Updated = "30.07.2024"

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

Write-host ""
Write-host "Information" -ForegroundColor Yellow
Write-host "This script will prompt for the Name and Email to the Shared Mailbox"
Write-host "Then open a CSV where you will have to enter their firstname and lastname"
Write-host ""
Write-host "Remember to Save the notepad document with Ctrl+S" -ForegroundColor Yellow
Write-host "If the user is not found it will ask you to manually input the email to the user"
Write-host ""

# Gather Information
$SharedMailboxName = Read-Host "Enter the Name of the Shared Mailbox"
$SharedMailbox = Read-Host "Enter the Shared Mailbox Email Address"
Write-host ""

# Extract the part before the "@" symbol
$SharedMailboxAlias = $SharedMailbox -split "@" | Select-Object -First 1

# Define the path to the CSV file
$csvPath = "$env:TEMP\users.csv"

# Create the CSV content
$csvContent = @"
Name
Firstname Lastname
"@ 

# Write the CSV content to a file
$csvContent | Out-File -FilePath $csvPath -Encoding utf8

# Open Notepad with the file
Start-Process "notepad.exe" -ArgumentList $csvPath

# Create Shared Mailbox and run Configuration
New-Mailbox -Shared -Name "$SharedMailboxName" -PrimarySmtpAddress "$SharedMailbox" -Alias "$SharedMailboxAlias"
Write-host ""
Write-host "Pending 10 Seconds before config running"
Timeout /T 10
Set-Mailbox -Identity "$SharedMailbox" -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true
Write-host ""

# Wait until users.csv is saved
Write-host "Please update the CSV file with the names of the users and Save it before clicking Enter"
Write-host ""
Pause
Write-host ""

# Read the CSV file
$users = Import-Csv -Path $csvPath

# Loop through each user in the CSV
foreach ($user in $users) {
    $name = $user.Name

    # Search for the user in Exchange Online
    $recipient = Get-Recipient -Filter "Name -like '*$name*'"
    
    if ($recipient -ne $null) {
        Write-Host "User with name '$name' found."

        # Confirm access before adding user to the shared mailbox
        $confirmation = Read-Host "Do you want to grant access to $name? (Y/N)"

        if ($confirmation -eq "Y" -or $confirmation -eq "y") {
            # Get the email address of the user
            $email = $recipient.PrimarySmtpAddress

            # Add user to the shared mailbox
            Add-MailboxPermission -Identity $SharedMailbox -User "$email" -AccessRights FullAccess -Confirm:$false
            Write-Host "Access granted to $name ($email)."

            Add-RecipientPermission $SharedMailbox -AccessRights SendAs -Trustee "$email" -Confirm:$false
            Write-Host "Granted SendAs Access to $user on $sharedMailbox"
            Write-host ""

        } else {
            Write-Host "Access not granted to $name."
            Write-host ""
        }
    } else {
        Write-Host "User with name '$name' not found in Exchange Online."
        
        # Prompt for email address
        $email = Read-Host "Enter the email address for $name"
        
        # Add user to the shared mailbox with the provided email address
        Add-MailboxPermission -Identity $SharedMailbox -User "$email" -AccessRights FullAccess -Confirm:$false
        Write-Host "Access granted to $name ($email)."
        
        # Add SendAs permission to the user on the shared mailbox with the provided email address
        Add-RecipientPermission $SharedMailbox -AccessRights SendAs -Trustee "$email" -Confirm:$false
        Write-Host "Granted SendAs Access to $email on $sharedMailbox"
        Write-host ""
    }
}

# Define the GiveAccess function
function GiveAccess {
    param (
        [string]$email,
        [string]$access = "FullAccess"
    )

    $errors = @()  # Initialize an empty array to store errors

    # Define the commands based on the value of $access
    switch ($access) {
        "FullAccess" {
            # $SharedMailbox = Read-Host "Enter the Shared Mailbox Email Address"
            Add-MailboxPermission -Identity $SharedMailbox -User $email -AccessRights FullAccess -ErrorVariable FullAccessError -ErrorAction SilentlyContinue
            Add-RecipientPermission -Identity $SharedMailbox -AccessRights SendAs -Trustee $email -Confirm:$false -ErrorVariable SendAsError -ErrorAction SilentlyContinue

            if ($FullAccessError) {
                $errors += $FullAccessError.Exception.Message
            }

            if ($SendAsError) {
                $errors += $SendAsError.Exception.Message
            }

            if (-not $FullAccessError -and -not $SendAsError) {
                $GiveAccessMessage = "Granting SendAs and Full access for $email on $SharedMailbox."
            }
        }
        "SendAs" {
            # $SharedMailbox = Read-Host "Enter the Shared Mailbox Email Address"
            Add-RecipientPermission -Identity $SharedMailbox -AccessRights SendAs -Trustee $email -Confirm:$false -ErrorVariable SendAsError -ErrorAction SilentlyContinue

            if ($SendAsError) {
                $errors += $SendAsError.Exception.Message
            }

            if (-not $SendAsError) {
                $GiveAccessMessage = "Granting SendAs Access for $email on $SharedMailbox."
            }
        }
        Default {
            Write-Host "Invalid access type."
            return
        }
    }

    # Output errors
    if ($errors.Count -gt 0) {
        Write-Host "Errors occurred:"
        foreach ($error in $errors) {
            Write-Host $error -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "$GiveAccessMessage" -ForegroundColor Green
    }
}

Write-host ""
Write-host ""
Write-host ""
Write-host ""
Write-host "If there is some issues you can add users manually to this Shared Mailbox with the command GiveAccess user@domain.com" -ForegroundColor Yellow

} else {
Write-host ""
Write-host "If you have used Connect Customer PowerShell you can connect with:"
Write-host "Connect Exchange `"Customer`""
    Return
}