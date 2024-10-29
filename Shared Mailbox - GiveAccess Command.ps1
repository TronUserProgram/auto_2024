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

Write-host ""
Write-host "Once you are connected to the relevant customer with correct credetials"
Write-host ""
Write-host "Connect-ExchangeOnline -UserPrincipalName "firstname.lastname@cegal.onmicrosoft.com" -DelegatedOrganization "CUSTOMER.onmicrosoft.com""
Write-host "Connect-ExchangeOnline -UserPrincipalName "365ADMINUSER@CUSTOMER.onmicrosoft.com""
Write-host ""
Write-host "You can then use the command GiveAccess user@company.com"
Write-host "You may also add SendAs or FullAccess - GiveAccess user@company.com **"
