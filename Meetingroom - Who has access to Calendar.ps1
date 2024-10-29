# Developed by Stian Kvia - Customer Support UA

# This Powershell will check the access of a Meetingroom or a Users calendar if u enter that instead.

# Gather Information
$Meetingroom = Read-Host "Enter the Meetingroom Email Address"

# Replace 'RoomName' with the actual name of your meeting room
$roomCalendar = Get-MailboxFolderPermission -Identity ${Meetingroom}:\Calendar

# Check permissions
if ($roomCalendar -ne $null) {
    Write-Output "The following users/groups can book the meeting room: $Meetingroom"
    $roomCalendar | Select-Object User, AccessRights
} else {
    Write-Output "No permissions found for booking the meeting room."
}