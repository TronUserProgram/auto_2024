# Prompt user to enter the meeting room email address
$meetingRoom = Read-Host "Enter the meeting room email address"

# Get the room lists
$roomLists = Get-DistributionGroup -RecipientTypeDetails RoomList

# Display list of meeting rooms with numbers
Write-Host "Room Lists:"
for ($i = 0; $i -lt $roomLists.Count; $i++) {
    Write-Host "$($i + 1). $($roomLists[$i].Name) - $($roomLists[$i].PrimarySmtpAddress)"
}

# Prompt user to select a room list
$selectedRoomListIndex = Read-Host "Enter the number corresponding to the room list you want to add $meetingRoom to"

# Validate user input
if ($selectedRoomListIndex -ge 1 -and $selectedRoomListIndex -le $roomLists.Count) {
    $selectedRoomList = $roomLists[$selectedRoomListIndex - 1]
    Write-Host "You selected: $($selectedRoomList.Name) - $($selectedRoomList.PrimarySmtpAddress)"
    
    # Add the meeting room to the selected room list
    Try {
    Add-DistributionGroupMember -Identity $selectedRoomList.Identity -Member $meetingRoom
    Write-Host "$meetingRoom added to $($selectedRoomList.Name) Roomlist"
    } catch {
    Write-Host "There was an error adding to Roomlist, make sure the resource is a room"
    }
    
} else {
    Write-Host "Invalid selection. Please enter a number between 1 and $($roomLists.Count)"
}
