# Developed by Stian Kvia
$Updated = "27.10.2024"

# Pending to Create Powershell Modules that are called from Connect Customer Directory instead of everything inside one PowerShell

# This sets a variable that the Connect Customer is Enabled for Scripts that run modules from this PowerShell
$Connect_Customer = "Enabled"
$SecretURL = "https://company.secretservercloud.eu/app/#/secrets/$AzureSecretServer/general"

function Connect {
    param (
        [string]$access = "Exchange",
        [string]$customer = "",
        [string]$user = $env:USERNAME,
        [string]$siteUrl = ""  # Add a parameter for the site URL
    )

    # Define email addresses based on the name
    switch ($User.ToLower()) {
        "stian" { $Agent = "admstiankv@comany.com"; $Admin = "stian.kvia@company.onmicrosoft.com"; break }
        "stiankv" { $Agent = "admstiankv@company.com"; $Admin = "stian.kvia@company.onmicrosoft.com"; break }
        default {
            Write-Host "Invalid name entered. Please enter Stian, Monica, Bea or Christoffer." -ForegroundColor Red
            exit
        }
    }


# Extract the username part from $Admin (firstname.lastname)
    $username = $Admin.Split('@')[0]
    $userfirstname = $username.Split('.')[0]
    $userlastname = $username.Split('.')[1]


# Construct the new personal email address using (firstname.lastname@)
    $Company1Admin = "la.$username@1st_customerdomain.onmicrosoft.com"
    $Company2Admin = "$username@2nd_customerdomain.onmicrosoft.com"


# Define a hashtable to store mappings between company names and their details
# ConnectionType = "simple" should be used if the tenant uses their own admin
# Following values are used by Connect Exchange: Url, Admin, SecretServer, ConnectionType, Message
# Following values are used by Connect Azure: Url, AzureAdmin, AzureSecretServer, AzureMessage
$Companies = @{
    "template" = @{
        "Url" = "customer.onmicrosoft.com"
        
        "Admin" = "admin@customer.onmicrosoft.com"
        "SecretServer" = "0000"
        "ConnectionType" = "simple"
        "Message" = "Remember to PIM the account"
        
        "AzureAdmin" = "admin@customer.onmicrosoft.com"
        "AzureSecretServer" = "0000"
        "AzureMessage" = "Remember to PIM the account"
    }
    "company" = @{
        "Url" = "company.onmicrosoft.com"
        "Admin" = $Agent
        "ConnectionType" = "simple"
        "Message" = "Remember to PIM the Account"
        "AzureMessage" = "Remember to PIM the account"
    }
    "force" = @{
        "ForceConnect" = "True"
    }
}


    # Find the company based on input
    $matchingCompanies = $Companies.GetEnumerator() | Where-Object { $_.Key.ToLower().Replace(" ", "").Contains($Customer.ToLower().Replace(" ", "").Substring(0, [Math]::Min(5, $Customer.Length))) }

    if ($matchingCompanies.Count -eq 0) {
        Write-Host "No matching company found." -ForegroundColor Red
        return
    }

    # If there's more than one matching company, select the first one
    if ($matchingCompanies.Count -gt 1) {
        $selectedCompany = $matchingCompanies[0].Value
    } else {
        $selectedCompany = $matchingCompanies.Value
    }


       # Set company details
    $Company = $selectedCompany.Url

    # Check if Admin is defined in the switch statement, otherwise use the one from $selectedCompany
   # if (-not $Admin) {
   #     $Admin = $selectedCompany.Admin
   # }

   # Check if Admin is defined in the switch statement, otherwise use the default admin
    if (-not $selectedCompany.Admin) {
        $Admin = "$Admin"
    } else {
        $Admin = $selectedCompany.Admin
    }

   # Check if Azure Admin is defined in the switch statement, otherwise use the default admin
    if (-not $selectedCompany.AzureAdmin) {
        $AzureAdmin = "$Admin"
    } else {
        $AzureAdmin = $selectedCompany.AzureAdmin
    }

    # Check if ConnectionType is defined in the switch statement, otherwise use the one from $selectedCompany
    if (-not $selectedCompany.ConnectionType) {
        $ConnectionType = "default"  # Set a default value
    } else {
        $ConnectionType = $selectedCompany.ConnectionType
    }

    $SecretServer = $selectedCompany.SecretServer
    $AzureSecretServer = $selectedCompany.AzureSecretServer
    $Message = $selectedCompany.Message
    $AzureMessage = $selectedCompany.AzureMessage
    $ForceConnect = $selectedCompany.ForceConnect

    # Define the command based on the value of $access
    switch ($access) {
        "Exchange" {
            if ($ConnectionType.ToLower() -eq "simple") {
$command = @"
Connect-ExchangeOnline -UserPrincipalName '$Admin'
Connected_Info
"@
            } else {
$command = @"
Connect-ExchangeOnline -UserPrincipalName '$Admin' -DelegatedOrganization '$Company'
Connected_Info
"@
            }
            Write-Host "Connecting Exchange towards $Company with Admin: $Admin"
            Write-host "$Message" -ForegroundColor Green
            # Check if SecretServer value is present
            if ($SecretServer) {
              Set-Clipboard -Value $SecretURL
              Write-Host "SecretServer URL Copied to Clipboard" -ForegroundColor Yellow
            }
        }
        "Test" {
            if ($ConnectionType.ToLower() -eq "simple") {
                $command = "Connect-ExchangeOnline -UserPrincipalName '$Admin'"
            } else {
                $command = "Connect-ExchangeOnline -UseRPSSession -UserPrincipalName '$Admin' -DelegatedOrganization '$Company'"
            }
            Write-Host "Connecting Exchange towards $Company with Admin: $Admin"
            Write-host "$Message" -ForegroundColor Green
            # Check if SecretServer value is present
            if ($SecretServer) {
              Set-Clipboard -Value $SecretURL
              Write-Host "SecretServer URL Copied to Clipboard" -ForegroundColor Yellow
            }
        }
        "ExchangeOut" {
            if ($ConnectionType.ToLower() -eq "simple") {
                Write-host = "Connect-ExchangeOnline -UserPrincipalName '$Admin'"
            } else {
                Write-host = "Connect-ExchangeOnline -UseRPSSession -UserPrincipalName '$Admin' -DelegatedOrganization '$Company'"
            }
            Write-Host "Connecting Exchange towards $Company with Admin: $Admin"
            Write-host "$Message" -ForegroundColor Green
            # Check if SecretServer value is present
            if ($SecretServer) {
              Set-Clipboard -Value $SecretURL
              Write-Host "SecretServer URL Copied to Clipboard" -ForegroundColor Yellow
            }
        }
        "Azure" {
            if ($ConnectionType.ToLower() -eq "simple") {
                $command = "Connect-AzureAD"
                
            } else {
                $command = "Connect-AzureAD"
            }
            Write-Host "To connect towards $Customer you need to login with the Admin: $AzureAdmin"
            Write-host "$AzureMessage" -ForegroundColor Green
            if ($AzureSecretServer) {
              Set-Clipboard -Value $SecretURL
              Write-Host "SecretServer URL Copied to Clipboard" -ForegroundColor Yellow
            }
        }
        "AzAzure" {
            if ($ConnectionType.ToLower() -eq "simple") {
                $command = "Connect-AzAccount"
                
            } else {
                $command = "Connect-AzAccount"
            }
            Write-Host "To connect towards $Customer you need to login with the Admin: $AzureAdmin"
            Write-host "$AzureMessage" -ForegroundColor Green
            if ($AzureSecretServer) {
              Set-Clipboard -Value $SecretURL
              Write-Host "SecretServer URL Copied to Clipboard" -ForegroundColor Yellow
            }
        }
        "Msol" {
            if ($ConnectionType.ToLower() -eq "simple") {
                $command = "Connect-MsolService"
                
            } else {
                $command = "Connect-MsolService"
            }
            Write-Host "To connect towards $Customer you need to login with the Admin: $AzureAdmin"
            Write-host "$AzureMessage" -ForegroundColor Green
            if ($AzureSecretServer) {
              Set-Clipboard -Value $SecretURL
              Write-Host "SecretServer URL Copied to Clipboard" -ForegroundColor Yellow
            }
        }
        "Teams" {
                # Ensure siteUrl is provided
                if (-not $siteUrl) {
                    $siteUrl = Read-Host "Please enter the URL to the site"
                    return
                }

                # Default command without forced authentication
                $command = "Connect-PnPOnline -Url $siteUrl -Interactive"

                # Check if ForceAuthentication is requested
                if ($ForceConnect.ToLower() -eq "true") {
                    $command = "Connect-PnPOnline -Url $siteUrl -Interactive"
                    Write-Host "Connecting to Teams site: $siteUrl with forced authentication."
                } else {
                    $command = "Connect-PnPOnline -Url $siteUrl -Interactive -ForceAuthentication"
                    Write-Host "Connecting to Teams site: $siteUrl."
                }
        }
        "Info" {
                
                Write-Host "Exchange Online Information" -ForegroundColor Green
                Write-Host "Use the account: $Admin"
                if ($SecretServer) {
                    Write-Host "Secret Server: $SecretURL"
                }
                if ($Message) {
                    Write-Host "Information: $Message" -ForegroundColor Yellow
                }
                Write-Host ""
            # Check if SecretServer value is present
                Write-Host "Azure / Entra Information" -ForegroundColor Green
                if ($AzureAdmin) {
                    Write-Host "Use the account: $AzureAdmin"
                }
                if ($AzureSecretServer) {
                    Write-Host "Secret Server: $SecretURL"
                }
                if ($AzureMessage) {
                    Write-Host "Information: $AzureMessage" -ForegroundColor Yellow
                }
            
        }

        Default {
            Write-Host "Invalid access type."
            return
        }
    }

    # Execute the command using Invoke-Expression
                try {
                Invoke-Expression $command
                }
            catch
                {
                }
 #   Write-host $command
}

Write-Host @"

                                                             *Developed by Stian Kvia
   ______                            __     ______            __    __         __
  / ____/___  ____  ____  ___  _____/ /_   / ____/___  ____ _/ /_  / /__  ____/ /
 / /   / __ \/ __ \/ __ \/ _ \/ ___/ __/  / __/ / __ \/ __ '/ __ \/ / _ \/ __  / 
/ /___/ /_/ / / / / / / /  __/ /__/ /_   / /___/ / / / /_/ / /_/ / /  __/ /_/ /  
\____/\____/_/ /_/_/ /_/\___/\___/\__/  /_____/_/ /_/\__,_/_.___/_/\___/\__,_/  
Updated: $Updated
                                                                           
"@ -ForegroundColor Green

Write-host ""
Write-host "You may now use the following command to connect to our customers"
Write-host ""
Write-host "Connect Exchange `"Customer`""
Write-host ""
Write-host "Write help to get more commands"


######################################################################### END OF FUNCTION: CONNECT #################################################

# Configure Teams

function Configure {
    param (
        [string]$component,
        [string]$company
    )

    # Ensure the component and company are provided
    if (-not $component -or -not $company) {
        Write-Host "Both component and company parameters are required." -ForegroundColor Red
        return
    }

    # Ensure the component is "Teams"
    if ($component.ToLower() -ne "teams") {
        Write-Host "Invalid component. This function only supports 'Teams'." -ForegroundColor Red
        return
    }

    # Validate the company (you can add more validation as needed)
    if ($company.ToLower() -ne "sval") {
        Write-Host "Invalid company. This function only supports 'Sval'." -ForegroundColor Red
        return
    }

    # Execute the commands
    try {
        Set-PnPSite -DisableCompanyWideSharingLinks Disabled
        Write-Host "Disabled Company Wide Sharing Links" -ForegroundColor Green
        } catch {
        Write-Host "Failed to configure Teams for $company. | CompanyWideSharingLinks" -ForegroundColor Red
        }
    try {
        Set-PnPSite -DefaultLinkToExistingAccess:$true
        Write-Host "Default to People with existing access" -ForegroundColor Green
        } catch {
        Write-Host "Failed to configure Teams for $company. | LinkToExistingAccess" -ForegroundColor Red
        }
        
}

######################################################################### END OF FUNCTION: CONFIGURE TEAMS #################################################

# Define the GiveAccess function
function GiveAccess {
    param (
        [string]$email,
        [string]$access = "FullAccess"
    )

    $errors = @()  # Initialize an empty array to store errors

    # Define the commands based on the value of $access
    switch ($access) {
        "FullAccess" {
            $SharedMailbox = Read-Host "Enter the Shared Mailbox Email Address"
            Add-MailboxPermission -Identity $SharedMailbox -User $email -AccessRights FullAccess -ErrorVariable FullAccessError -ErrorAction SilentlyContinue
            Add-RecipientPermission -Identity $SharedMailbox -AccessRights SendAs -Trustee $email -Confirm:$false -ErrorVariable SendAsError -ErrorAction SilentlyContinue

            if ($FullAccessError) {
                $errors += $FullAccessError.Exception.Message
            }

            if ($SendAsError) {
                $errors += $SendAsError.Exception.Message
            }

            if (-not $FullAccessError -and -not $SendAsError) {
                $GiveAccessMessage = "Granting SendAs and Full access for $email on $SharedMailbox."
            }
        }
        "SendAs" {
            $SharedMailbox = Read-Host "Enter the Shared Mailbox Email Address"
            Add-RecipientPermission -Identity $SharedMailbox -AccessRights SendAs -Trustee $email -Confirm:$false -ErrorVariable SendAsError -ErrorAction SilentlyContinue

            if ($SendAsError) {
                $errors += $SendAsError.Exception.Message
            }

            if (-not $SendAsError) {
                $GiveAccessMessage = "Granting SendAs Access for $email on $SharedMailbox."
            }
        }
        Default {
            Write-Host "Invalid access type."
            return
        }
    }

    # Output errors
    if ($errors.Count -gt 0) {
        Write-Host "Errors occurred:"
        foreach ($error in $errors) {
            Write-Host $error -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "$GiveAccessMessage" -ForegroundColor Green
    }
}

######################################################################### END OF FUNCTION: GIVEACCESS #################################################
# Module: Connected
# Get Exchange Online Connected Customer
function Connected {
    try {
        $acceptedDomains = Get-AcceptedDomain | Where-Object { $_.Default -eq $true }

        if ($acceptedDomains) {
            Write-Host ""
            Write-Host "You are currently connected to the domain" -NoNewline
            Write-Host " $($acceptedDomains.Name) " -ForegroundColor Yellow
        } else {
            Write-Host "You are not currently connected to any default domain." -ForegroundColor Red
            Write-Host "You have to connect with Connect Exchange 'Customer'"
        }
    }
    catch {
        Write-Host "Unable to retrieve accepted domains. Please ensure you are connected to Exchange Online."
        Write-Host "You have to connect with Connect Exchange 'Customer'"
    }
}

######################################################################### END OF FUNCTION: CONNECTED #################################################
# Module: Connected_Info
# Get Exchange Online Connected Customer
function Connected_Info {
    try {
        $acceptedDomains = Get-AcceptedDomain | Where-Object { $_.Default -eq $true }

        if ($acceptedDomains) {
            Write-Host ""
            Write-Host "You are now connected to" -NoNewline
            Write-Host " $($acceptedDomains.Name) " -ForegroundColor Yellow
        } else {
            Write-Host "You are not currently connected to any default domain." -ForegroundColor Red
            Write-Host "You have to connect with Connect Exchange 'Customer'"
        }
    }
    catch {
        Write-Host "Unable to retrieve accepted domains. Please ensure you are connected to Exchange Online."
        Write-Host "You have to connect with Connect Exchange 'Customer'"
    }
}

######################################################################### END OF FUNCTION: CONNECTED #################################################
# Module: AutoConnect

function AutoConnect {

#try {
#    $acceptedDomains = Get-AcceptedDomain | Where-Object { $_.Default -eq $true } # Check if connected to the desired domain
#} catch {
#    Write-host "Not Connected to any Domain"
#    Write-host "Use the 'Connect Exchange `"Customer`" Command'"
#    return
#}
if ($acceptedDomains.Name -eq $domain) {
    Write-Host "You are currently connected to the domain" -NoNewline
    Write-Host " $($acceptedDomains.Name)" -ForegroundColor Yellow
} else {
    Write-Output "Connected to wrong tenant"
    try {
        Invoke-Expression $customer # Connect to the customer set in the script calling this function
    } catch {
    Write-Host "Cancelled" -ForegroundColor Red
    return
    }
}
}

######################################################################### END OF FUNCTION: AUTO CONNECT #################################################
function InstallModules {
# Check if connected to the desired domain
try {
    Write-host "Installing AzureAD"
    Install-Module AzureADPreview -AllowClobber -Force -Scope CurrentUser -Repository PSGallery
    Write-host "Installing ExchangeOnline"
    Install-Module ExchangeOnlineManagement -AllowClobber -Force -Scope CurrentUser
    Write-host "Installing MSOnline"
    Install-Module MSOnline -AllowClobber -Force -Scope CurrentUser
    Write-host "Installing Burn Toast"
    Install-Module -Name BurntToast -RequiredVersion 0.8.5 -AllowClobber -Force -Scope CurrentUser
    Write-host "Installing PNP Powershell"
    Install-Module -Name PnP.PowerShell -RequiredVersion 1.10.0 -Scope CurrentUser -Repository PSGallery -AllowClobber -Force
    Write-host "Installing Microsoft Teams"
    Install-Module -Name MicrosoftTeams -Scope CurrentUser -AllowClobber -Force

} catch {
    Write-host "Failed Install"
}

try {
    Write-host "Importing AzureAD"
    Import-Module AzureAdPreview -Force -Verbose
    Write-host "Importing ExchangeOnline"
    Import-Module ExchangeOnlineManagement -Force -Verbose
    Write-host "Importing MSOnline"
    Import-Module MSOnline -Force -Verbose
    Write-host "Importing PNP Powershell"
    Import-Module PnP.PowerShell -Force -Verbose
    Write-host "Importing Teams"
    Import-Module MicrosoftTeams -Force -Verbose
} catch {
    Write-host "Failed Import"
}
}


######################################################################### END OF FUNCTION: INSTALL MODULES #################################################


# Encrypter, Decrypter

# Encrypt or Decrypt Text from clipboard
function Crypt {
    $clipboardContent = Get-Clipboard

    if ($clipboardContent.StartsWith("RW5jcnlwdGVk")) {
        # Decryption
        $encodedText = $clipboardContent -replace "^RW5jcnlwdGVk", "" -replace "^\s+|\s+$"
        $decodedText = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($encodedText))
        $decodedText | Set-Clipboard
        Write-Host "Decrypted Text copied to clipboard."
    } else {
        # Encryption
        $encodedText = "RW5jcnlwdGVk" + [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($clipboardContent))
        $encodedText | Set-Clipboard
        Write-Host "Encrypted Text copied to clipboard."
    }
}

######################################################################### END OF FUNCTION: CRYPT #################################################

function Speak {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromRemainingArguments=$true, Mandatory=$true)]
        [string]$text
    )

    try {
        $TTS = New-Object -ComObject SAPI.SPVoice
        foreach ($voice in $TTS.GetVoices()) {
            if ($voice.GetDescription() -like "*- English*") {
                $TTS.Voice = $voice
                [void]$TTS.Speak($text)
                return
            }
        }
        throw "No English text-to-speech voice found - please install one."
    } catch {
        "⚠️ Error in line $($_.InvocationInfo.ScriptLineNumber): $($Error[0])"
    }
}

######################################################################### END OF FUNCTION: SPEAK #################################################

# Gives a strikethrough on all text in clipboard
function Strike {
    # Get text from clipboard
    $text = Get-Clipboard

    $strikethrough = $text -replace ".","$($([char]0x0336))$&"

    # Copy strikethrough text back to clipboard
    $strikethrough | Set-Clipboard

    Write-Host "Text with strikethrough effect copied to clipboard:"
}

######################################################################### END OF FUNCTION: STRIKE #################################################

function Create {
    param(
        [string]$SharedMailbox,
        [string]$Company
    )

    # Get the directory of the currently executing script
    $scriptDirectory = $PSScriptRoot

    # Determine the script file name based on whether $Company is provided
    if (-not [string]::IsNullOrEmpty($Company)) {
        $scriptFileName = "\CreateMailbox\$Company - create shared mailbox.ps1"
    } else {
        $scriptFileName = "Shared Mailbox - Create Mailbox and Add Multiple Users via Email.ps1"
    }

    # Construct the full path to the PowerShell script file
    $scriptPath = Join-Path -Path $scriptDirectory -ChildPath $scriptFileName

    # Check if the script file exists
    if (Test-Path $scriptPath) {
        # Call the script file with the provided parameter
        & $scriptPath -SharedMailbox $SharedMailbox
    } else {
        Write-Host "Error: Script file not found at $scriptPath"
    }
}
######################################################################### END OF FUNCTION: CREATE #################################################

# Create Random Password
Function Get-RandomPassword {
    # Define parameters
    param(
        [Parameter(ValueFromPipeline=$false)]
        [ValidateRange(1,256)]
        [int]$PasswordLength = 14
    )
 
    # ASCII Character set for Password
    $CharacterSet = @{
        Lowercase   = (97..122 -ne 108) | Get-Random -Count 10 | % {[char]$_} # Exclude 'l' (108)
        Uppercase   = (65..90 -ne 79 -ne 73 -ne 83 -ne 66 -ne 71) | Get-Random -Count 10 | % {[char]$_} # Exclude 'O' (79), 'I' (73), 'S' (83), 'B' (66), 'G' (71)
        Numeric     = (48..57 -ne 48 -ne 49 -ne 53 -ne 56 -ne 54) | Get-Random -Count 10 | % {[char]$_} # Exclude '0' (48), '1' (49), '5' (53), '8' (56), '6' (54)
        SpecialChar = 33, 37, 42, 35 | Get-Random -Count 4 | % {[char]$_} # Include specific special characters '! % * #'
    }
 
    # Frame Random Password from given character set
    $StringSet = $CharacterSet.Uppercase + $CharacterSet.Lowercase + $CharacterSet.Numeric + $CharacterSet.SpecialChar
 
    $password = -join (Get-Random -Count $PasswordLength -InputObject $StringSet)
    
    # Copy password to clipboard

    $password | Set-Clipboard
    Write-output $password
    Write-host "Password copied to clipboard."
}

# Create Random Password
Function Get-Password {
    # Define parameters
    param(
        [Parameter(ValueFromPipeline=$false)]
        [ValidateRange(1,256)]
        [int]$PasswordLength = 14
    )
 
    # ASCII Character set for Password
    $CharacterSet = @{
        Lowercase   = (97..122 -ne 108) | Get-Random -Count 10 | % {[char]$_} # Exclude 'l' (108)
        Uppercase   = (65..90 -ne 79 -ne 73 -ne 83 -ne 66 -ne 71) | Get-Random -Count 10 | % {[char]$_} # Exclude 'O' (79), 'I' (73), 'S' (83), 'B' (66), 'G' (71)
        Numeric     = (48..57 -ne 48 -ne 49 -ne 53 -ne 56 -ne 54) | Get-Random -Count 10 | % {[char]$_} # Exclude '0' (48), '1' (49), '5' (53), '8' (56), '6' (54)
        SpecialChar = 33, 37, 42, 35 | Get-Random -Count 4 | % {[char]$_} # Include specific special characters '! % * #'
    }
 
    # Frame Random Password from given character set
    $StringSet = $CharacterSet.Uppercase + $CharacterSet.Lowercase + $CharacterSet.Numeric + $CharacterSet.SpecialChar
 
    $password = -join (Get-Random -Count $PasswordLength -InputObject $StringSet)
    
    # Copy password to clipboard

    $password | Set-Clipboard
    Write-output $password
    Write-host "Password copied to clipboard."
}

######################################################################### END OF FUNCTION: RandomPassword #################################################

# Wait

function Wait {
    param (
        [int]$minTime = 50,
        [int]$maxTime = 90
    )

    $waitTime = Get-Random -Minimum $minTime -Maximum ($maxTime + 1)
    Write-Host "Pending $waitTime seconds"
    Start-Sleep -Seconds $waitTime
}

######################################################################### END OF FUNCTION: Wait #################################################

# Hidden Help

# This provides Help Text
function HiddenHelp {

Write-Host @"
Hidden
    __  __     __    
   / / / /__  / /___ 
  / /_/ / _ \/ / __ \
 / __  /  __/ / /_/ /
/_/ /_/\___/_/ .___/ 
            /_/      
"@ -ForegroundColor Green

Write-host @"

With Connect Customer PowerShell you can make the following commands

Crypt                         | This will encrypt or decrypt any text in your clipboard
                              |
Speak                         | This will speak the text you enter after Speak
                              |
Strike                        | This will strike out any text in your clipboard
                              |
InstallModules                | This will Install all PowerShell Modules
                              |
Get-RandomPassword 12         | This will give you a random password of 12 characters
                              |
Wait 10 20                    | Waits between 10 - 20 seconds until next operation, Useful for when its needed

"@ -ForegroundColor White


}

# This provides Help Text
function Help {

Write-Host @"

    __  __     __    
   / / / /__  / /___ 
  / /_/ / _ \/ / __ \
 / __  /  __/ / /_/ /
/_/ /_/\___/_/ .___/ 
            /_/      
"@ -ForegroundColor Green

Write-host @"

With Connect Customer PowerShell you can make the following commands

Connect Exchange `"Customer`"   | This will Connect towards the customer with Exchange Online
                              |
Connect Azure `"Customer`"      | This will Connect towards the customer with Azure Connect
                              |
Connect Info `"Customer`"       | This will Give you information about what account to use
                              |
Connect ExOp `"Customer`"       | This will Connect towards the customer with Exchange Online
                              |
Connect Teams *               | This will Connect towards Sharepoint of Customer * FORCE will make it force Reconnect
                              |
Connected                     | This will show you the current connected customer
                              |
GiveAccess user@domain.com    | This will add the user to a shared mailbox with full access on the customer domain
                              |
Create SharedMailbox          | This will let you create a shared mailbox and add users with full access in format mail@domain.com
                              |
Crypt                         | This will encrypt or decrypt any text in your clipboard
                              |
Strike                        | This will strike out any text in your clipboard
                              |
InstallModules                | This will Install all PowerShell Modules
                              |
Get-RandomPassword 12         | This will give you a random password of 12 characters

"@ -ForegroundColor White


}

######################################################################### END OF FUNCTION: HELP #################################################
