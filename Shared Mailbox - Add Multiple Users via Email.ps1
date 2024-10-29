# Developed by Stian Kvia - Customer Support UA

# Gather Information
$SharedMailbox = Read-Host "Enter the Shared Email Address"
Write-host ""

# Define the path to the CSV file
$csvPath = "$env:TEMP\users.csv"


# Create the CSV content
$csvContent = @"
email
user@domain.com
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
Write-host "Please update the CSV file with the emails of the users and Save it before clicking Enter"
Write-host ""
Pause
Write-host ""

# Read the CSV file
$emails = Import-Csv -Path $csvPath

# Grant FullAccess and SendAs permissions for each email address
foreach ($email in $emails) {
    $user = $email.email

    # Grant FullAccess permission
    Add-MailboxPermission -Identity $SharedMailbox -User $user -AccessRights FullAccess
    Write-Host "Granted FullAccess to $user on $SharedMailbox"

    # Grant SendAs permission
    Add-RecipientPermission $SharedMailbox -AccessRights SendAs -Trustee $user
    Write-Host "Granted SendAs Access to $user on $SharedMailbox"
    Write-Host ""
}