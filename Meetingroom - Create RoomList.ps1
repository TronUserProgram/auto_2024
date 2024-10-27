# Developed by Stian Kvia
$Updated = "27.10.2024"

# Following Commands are used
# New-DistributionGroup -Name "$RoomListName" -Alias "$RoomListAlias" -RoomList

# Beginning of Script
Write-host "Beginning of PowerShell Script" -ForegroundColor Green
Write-Host ""

$RoomListName = Read-Host "Enter the Name of the Roomlist, Usually the Location"

# Automatically create the alias by replacing spaces with periods
$RoomListAlias = $RoomListName -replace " ", "."

Try {
    # Attempt to create the room list
    New-DistributionGroup -Name "$RoomListName" -Alias "$RoomListAlias" -RoomList
    Write-Host "Room list '$RoomListName' with alias '$RoomListAlias' created successfully." -ForegroundColor Green
} Catch {
    # Display the error message if something goes wrong
    Write-Host "There was an error creating the room list:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
}
