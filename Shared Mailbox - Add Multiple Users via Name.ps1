# Developed by Stian Kvia - Customer Support UA

Clear-host
Write-host "This script will prompt for the Email to the Shared Mailbox and open a CSV where you will have to enter their firstname and lastname"
Write-host "If the user is not found it will ask you to manually input the email to the user"
Write-host ""
Timeout /T 3
Write-host ""

# Gather Information
$SharedMailbox = Read-Host "Enter the Shared Email Address"
Write-host ""

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

# Connect to Exchange Online
Write-host "Make sure you are connected to the correct customer"
Timeout /T 5
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
            Add-MailboxPermission -Identity $SharedMailbox -User $email -AccessRights FullAccess
            Write-Host "Access granted to $name ($email)."

            Add-RecipientPermission $SharedMailbox -AccessRights SendAs -Trustee $email
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
        Add-MailboxPermission -Identity $SharedMailbox -User $email -AccessRights FullAccess
        Write-Host "Access granted to $name ($email)."
        
        # Add SendAs permission to the user on the shared mailbox with the provided email address
        Add-RecipientPermission $SharedMailbox -AccessRights SendAs -Trustee $email
        Write-Host "Granted SendAs Access to $email on $sharedMailbox"
        Write-host ""
    }
}