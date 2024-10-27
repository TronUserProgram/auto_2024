# Customer Connect and PowerShell Scripts for Managing Exchange Online, Azure, and More

### Description
`CustomerConnect.ps1` provides quick connections to customers with predefined usernames for multiple users, along with various utility functions.

### Usage
```powershell
Connect-Exchange "Customer"       # Connect to the customer's Exchange Online
Connect-Azure "Customer"          # Connect to the customer's Azure
Connect-Info "Customer"           # Display information about the account to use
Connect-ExOp "Customer"           # Connect to the customer's Exchange Online
Connect-Teams *                   # Connect to customer's SharePoint (use FORCE to reconnect)
Connected                         # Display the currently connected customer
GiveAccess user@domain.com        # Grant full access to a shared mailbox for a user on the customer's domain
Create-SharedMailbox              # Create a shared mailbox and grant full access to specified users
Crypt                             # Encrypt or decrypt text in the clipboard
Strike                            # Strike out text in the clipboard
InstallModules                    # Install all necessary PowerShell modules
Get-RandomPassword 12             # Generate a random password with 12 characters
