# Developed by Stian Kvia - Customer Support UA
$Updated = "31.07.2024"

# Beginning of Script
Write-host "Beginning of PowerShell Script" -ForegroundColor Green
Write-Host ""

########################### Check Module and AutoConnect START
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

########################### Check Module and AutoConnect END

# Gather Information
$SharedMailbox = Read-Host "Enter the Shared Email Address"

try {
Set-Mailbox -Identity "$SharedMailbox" -MessageCopyForSentAsEnabled $true -MessageCopyForSendOnBehalfEnabled $true
} catch {
Write-Host "Failed to configure Shared Mailbox : $SharedMailbox $($_.Exception.Message)"
}

# Confirmation Configured Shared Mailbox
Write-Host ""
Write-Host "Shared mailbox configuration completed successfully."

} else {
Write-host ""
Write-host "If you have used Connect Customer PowerShell you can connect with:"
Write-host "Connect Exchange `"Customer`""
    Return
}