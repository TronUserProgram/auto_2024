# Developed by Stian Kvia - Customer Support UA
# Verson 1.030624

# Following Commands are used
# Get-CalendarProcessing -Identity "EMAIL"
# Set-CalendarProcessing -Identity "EMAIL"

# This will gather the Calendar Processing of a Meetingroom so you easily
# can transfer Settings from and old meetingroom to a new one
# The command is available in clipboard after

# Beginning of Script
Write-host "Beginning of PowerShell Script" -ForegroundColor Green
Write-Host ""


# Check if $MeetingroomFrom is empty
Write-host "Set the Meetingroom to copy settings" -NoNewline
Write-host " From" -ForegroundColor Green

if ([string]::IsNullOrEmpty($MeetingroomFrom)) {
    $meetingRoomMessage = "Meeting room is not set, please enter a meetingroom email address"
} else {
    $meetingRoomMessage = "Current meeting room: $MeetingroomFrom Press Enter to keep or enter a new email"
}

# Display current value and prompt user to keep or enter a new one
$MeetingroomChoice = Read-Host -Prompt "$meetingRoomMessage"

# If user chooses to enter a new value
if ($MeetingroomChoice -ne "") {
    $MeetingroomFrom = $MeetingroomChoice
    Write-Host "Meeting room '$MeetingroomFrom' is set. Continuing with the script..." -ForegroundColor Green
} else {
    # If user chooses to keep the current value or presses Enter
        if ([string]::IsNullOrEmpty($MeetingroomFrom)) {
        Write-Host "Meeting room not provided. Exiting script."
        return
    }
    # Add the specific part of the script you want to execute when $MeetingroomFrom is set
    Write-Host "Meeting room '$MeetingroomFrom' is set. Continuing with the script..." -ForegroundColor Green
}


Write-host ""
Write-host "Set the Meetingroom to copy settings" -NoNewline
Write-host " To" -ForegroundColor Green

if ([string]::IsNullOrEmpty($MeetingroomTo)) {
    $meetingRoomMessage = "Meeting room is not set, please enter a meetingroom email address"
} else {
    $meetingRoomMessage = "Current meeting room: $MeetingroomTo Press Enter to keep or enter a new email"
}

# Display current value and prompt user to keep or enter a new one
$MeetingroomChoice = Read-Host -Prompt "$meetingRoomMessage"

# If user chooses to enter a new value
if ($MeetingroomChoice -ne "") {
    $MeetingroomTo = $MeetingroomChoice
    Write-Host "Meeting room '$MeetingroomTo' is set. Continuing with the script..." -ForegroundColor Green
} else {
    # If user chooses to keep the current value or presses Enter
        if ([string]::IsNullOrEmpty($MeetingroomTo)) {
        Write-Host "Meeting room not provided. Exiting script."
        return
    }
    # Add the specific part of the script you want to execute when $MeetingroomTo is set
    Write-Host "Meeting room '$MeetingroomTo' is set. Continuing with the script..." -ForegroundColor Green
}

    Write-host ""

    # MeetingroomFrom - Calendar Processing - Get calendar processing settings
    $calendarProcessing = Get-CalendarProcessing -Identity "$MeetingroomFrom"

    # MeetingroomFrom - Calendar Processing - Define a function to convert boolean values to PowerShell strings ('$true' or '$false')
    function ConvertToPowerShellBoolean($value) {
        if ($value) {
            return '$true'
       } else {
           return '$false'
       }
    }

    # MeetingroomFrom - Calendar Processing - Construct the Set-CalendarProcessing command
    $setCalendarProcessingCommand = "Set-CalendarProcessing -Identity ""$MeetingroomTo"" -AutomateProcessing $($calendarProcessing.AutomateProcessing) -AllowConflicts $(ConvertToPowerShellBoolean $calendarProcessing.AllowConflicts) -BookingType $($calendarProcessing.BookingType) -BookingWindowInDays $($calendarProcessing.BookingWindowInDays) -MaximumDurationInMinutes $($calendarProcessing.MaximumDurationInMinutes) -MinimumDurationInMinutes $($calendarProcessing.MinimumDurationInMinutes) -AllowRecurringMeetings $(ConvertToPowerShellBoolean $calendarProcessing.AllowRecurringMeetings) -EnforceCapacity $(ConvertToPowerShellBoolean $calendarProcessing.EnforceCapacity) -EnforceSchedulingHorizon $(ConvertToPowerShellBoolean $calendarProcessing.EnforceSchedulingHorizon) -ScheduleOnlyDuringWorkHours $(ConvertToPowerShellBoolean $calendarProcessing.ScheduleOnlyDuringWorkHours) -ConflictPercentageAllowed $($calendarProcessing.ConflictPercentageAllowed) -MaximumConflictInstances $($calendarProcessing.MaximumConflictInstances) -ForwardRequestsToDelegates $(ConvertToPowerShellBoolean $calendarProcessing.ForwardRequestsToDelegates) -DeleteAttachments $(ConvertToPowerShellBoolean $calendarProcessing.DeleteAttachments) -DeleteComments $(ConvertToPowerShellBoolean $calendarProcessing.DeleteComments) -RemovePrivateProperty $(ConvertToPowerShellBoolean $calendarProcessing.RemovePrivateProperty) -DeleteSubject $(ConvertToPowerShellBoolean $calendarProcessing.DeleteSubject) -AddOrganizerToSubject $(ConvertToPowerShellBoolean $calendarProcessing.AddOrganizerToSubject) -DeleteNonCalendarItems $(ConvertToPowerShellBoolean $calendarProcessing.DeleteNonCalendarItems) -TentativePendingApproval $(ConvertToPowerShellBoolean $calendarProcessing.TentativePendingApproval) -EnableResponseDetails $(ConvertToPowerShellBoolean $calendarProcessing.EnableResponseDetails) -OrganizerInfo $(ConvertToPowerShellBoolean $calendarProcessing.OrganizerInfo) -RequestOutOfPolicy $(ConvertToPowerShellBoolean $calendarProcessing.RequestOutOfPolicy) -AllRequestOutOfPolicy $(ConvertToPowerShellBoolean $calendarProcessing.AllRequestOutOfPolicy) -BookInPolicy $(ConvertToPowerShellBoolean $calendarProcessing.BookInPolicy) -AllBookInPolicy $(ConvertToPowerShellBoolean $calendarProcessing.AllBookInPolicy) -RequestInPolicy $(ConvertToPowerShellBoolean $calendarProcessing.RequestInPolicy) -AllRequestInPolicy $(ConvertToPowerShellBoolean $calendarProcessing.AllRequestInPolicy) -AddAdditionalResponse $(ConvertToPowerShellBoolean $calendarProcessing.AddAdditionalResponse) -AdditionalResponse '$($calendarProcessing.AdditionalResponse)' -RemoveOldMeetingMessages $(ConvertToPowerShellBoolean $calendarProcessing.RemoveOldMeetingMessages) -AddNewRequestsTentatively $(ConvertToPowerShellBoolean $calendarProcessing.AddNewRequestsTentatively) -ProcessExternalMeetingMessages $(ConvertToPowerShellBoolean $calendarProcessing.ProcessExternalMeetingMessages) -RemoveForwardedMeetingNotifications $(ConvertToPowerShellBoolean $calendarProcessing.RemoveForwardedMeetingNotifications) -RemoveCanceledMeetings $(ConvertToPowerShellBoolean $calendarProcessing.RemoveCanceledMeetings) -EnableAutoRelease $(ConvertToPowerShellBoolean $calendarProcessing.EnableAutoRelease) -PostReservationMaxClaimTimeInMinutes $($calendarProcessing.PostReservationMaxClaimTimeInMinutes)"

    # Output the command
    # Write-Output $setCalendarProcessingCommand
    # Set-Clipboard -Value "$setCalendarProcessingCommand"

    # MeetingroomFrom - Execute Commands to Configure new meetingroom

    try {
    Invoke-Expression $setCalendarProcessingCommand
    Write-Host "Configuration copied from $MeetingroomFrom to $MeetingroomTo" -ForegroundColor Green
} catch {
    if ($_.Exception.Message -like "*The specified mailbox Identity*doesn't exist*") {
        Write-Host ""
        Write-Host "The specified mailbox does not exist." -ForegroundColor Red
        Write-Host "$_" -ForegroundColor Yellow
    } else {
        Write-Host ""
        Write-Host "An unexpected error occurred:"
        Write-Host "$_"
    }
}