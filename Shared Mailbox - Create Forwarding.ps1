# Replace with your shared mailbox and forwarding email
$sharedMailbox = Read-Host "Enter the Shared Mailbox email"
$forwardToEmail = Read-Host "Enter the Forward to email"

# Enable email forwarding
Set-Mailbox -Identity $sharedMailbox -ForwardingSMTPAddress $forwardToEmail -DeliverToMailboxAndForward $true