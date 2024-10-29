# Developed by Stian Kvia - Customer Support UA

# This Powershell will check the access of a Meetingroom or a Users calendar if u enter that instead.
#
# Yes this is for users who have chosen to have their meeting room or outlook it in Norwegian, Lol

# Gather Information
$Meetingroom = Read-Host "Enter the Meetingroom Email Address"

# Replace 'RoomName' with the actual name of your meeting room
$roomCalendar = Get-MailboxFolderPermission -Identity ${Meetingroom}:\Kalender

# Check permissions
if ($roomCalendar -ne $null) {
    Write-Output "The following users/groups can book the meeting room: $Meetingroom"
    $roomCalendar | Select-Object User, AccessRights
} else {
    Write-Output "No permissions found for booking the meeting room."
}