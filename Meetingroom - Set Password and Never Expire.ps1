# Developed by Stian Kvia - Customer Support UA
$Updated = "23.07.2024"

# This Powershell is to Set a Password to a Meetingroom and set it to never expire.
# You will have to Manually use Connect-AzureAD -Identity "CUSTOMER365ADMIN@DOMAIN.COM"
# You can also use the UA Tool and Use 'Connect Azure Customer' note that not all customers are configured for Azure yet
# It will give you the Password in cleartext at the end.

########################### Check Module and Connected Customer

if ($Connect_Customer -eq "Enabled") {
} else {
    Write-host "Please Run Connect Customer PowerShell" -ForegroundColor Red
    return
}

#############################################################################################################

# Gather Information
$MeetingroomTo = Read-Host "Enter the Email Address to the Resource"
 
#Call the function to generate random password of 14 characters
$Password = Get-RandomPassword -PasswordLength 14
$PasswordSecure = ConvertTo-SecureString $Password -AsPlainText -Force

Set-AzureADUserPassword -ObjectId "$MeetingroomTo" -Password $PasswordSecure

Set-AzureADUser -ObjectID "$MeetingroomTo" -PasswordPolicies DisablePasswordExpiration

Write-Host "Password is: $Password"