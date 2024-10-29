# Remember to add a license for the mailbox to be generated
# This is necessary as the mailbox is required temporarily while converting from Mailbox to Meetingroom


# Prompt user to enter the meeting room email address
$meetingRoom = Read-Host "Enter the meeting room email address"

try {
    # Check if the mailbox exists, without showing errors
    $mailbox = Get-Mailbox -Identity $meetingRoom -ErrorAction Stop

    if ($mailbox) {
        # Try to set the mailbox type to Room, suppress errors
        Set-Mailbox -Identity $meetingRoom -Type Room -ErrorAction SilentlyContinue

        if ($?) {
            Write-Host "Success" -ForegroundColor Green
            Write-Host "Remember to remove the Exchange Online license now."
        } else {
            Write-Host "Failed to set mailbox type to Room for '$meetingRoom'." -ForegroundColor Yellow
            Write-Host "Ensure the Exchange Online license is applied during this process."
        }
    }
} catch {
    # Handle any errors (general exception catch)
    if ($_.Exception.Message -like "*couldn't be found*") {
        Write-Host "Error: The mailbox '$meetingRoom' could not be found." -ForegroundColor Red
        Write-Host "Ensure the email address is correct and that you temporarily have a license assigned." -ForegroundColor Yellow
        Write-Host "It has to minimum have an 'Exchange Online' License" -ForegroundColor Yellow
    } else {
        Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
    }
}
