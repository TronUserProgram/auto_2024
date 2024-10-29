# Developed by Stian Kvia - Customer Support UA
$Updated = "07.10.2024"

# Beginning of Script
Write-host "Beginning of PowerShell Script" -ForegroundColor Green

# Confirm Connected to Correct Customer
Connected
$confirmation = Read-Host "Are you sure you are connected to Exchange Online of the correct customer? (Y/N)"
if ($confirmation -eq "Y" -or $confirmation -eq "y") {
Write-Host ""
$group = Read-Host "Enter the Security Group Name"
Write-Host "Please wait, this may take some time."

# Loop through all shared mailboxes and check their permissions
Get-Mailbox -RecipientTypeDetails SharedMailbox | ForEach-Object {
    $mailbox = $_

    # Check Full Access permissions
    $fullAccessPermissions = Get-MailboxPermission -Identity $mailbox.Alias | Where-Object { $_.User -like "$group" }
    
    # Check Send As permissions
    $sendAsPermissions = Get-RecipientPermission -Identity $mailbox.Alias | Where-Object { $_.Trustee -like "$group" }
    
    # Output results if any permissions found
    if ($fullAccessPermissions -or $sendAsPermissions) {
        [pscustomobject]@{
            'Shared Mailbox          '  = $mailbox.DisplayName
            'Email          '           = $mailbox.PrimarySmtpAddress
            'FullAccess   '             = $fullAccessPermissions.AccessRights -join ', '  # FullAccess rights
            'SendAsAccess   '           = $sendAsPermissions.AccessRights -join ', '      # SendAs rights
        }
    }
}

} else {
Write-host ""
Write-host "If you have used Connect Customer PowerShell you can connect with:"
Write-host "Connect Exchange `"Customer`""
    Return
}