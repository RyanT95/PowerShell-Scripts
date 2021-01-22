#*******************************************
# Batch add users to shared inbox from CSV *
#*******************************************

# Check script is running as admin
#=============================================================
param([switch]$elevated)

function Is-Admin 
{
    # Get the current user, and check if they are an Administrator
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

# User does not have elevated privelages
if ((Is-Admin) -eq $false)
{
    if ($elevated) 
    {
        Write-Host "Script needs to be run as admin!" -ForegroundColor DarkRed
    } 
    else
	{
        # Start new elevated instance of PowerShell, user will be prompted with UAC login dialogue
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
	}
    
    # Close current instance of PowerShell
	exit
}

Write-Host "Running as admin...`n" -ForegroundColor Green
#=============================================================

# USER INPUT
$domain = Read-Host -Prompt 'Input company domain name (i.e. @domain.com)'

# Get user who ran script
$admCred = $env:UserName + $domain

# Connect to MS Exchange
try
{
	Import-Module ExchangeOnlineManagement -Verbose:$false
	Connect-ExchangeOnline -UserPrincipalName $admCred -ShowBanner:$false -ShowProgress:$true -Verbose:$false
}
catch
{
	Write-Host "$admCred is not a valid email! Check spelling." -ForegroundColor DarkRed
    exit
}

# CSV
$users = Import-Csv -Delimiter "," -Path "$PSScriptRoot\Contacts.csv"

# USER INPUT: shared inbox name
$sharedInboxName = Read-Host -Prompt 'Input full email address of shared inbox'

Write-Host "`n-----------------------------------------------------------------------------`n" -ForegroundColor DarkGray

# Error checking for valid shared inbox
try
{
    Get-Mailbox -RecipientTypeDetails SharedMailbox -Identity $sharedInboxName -ErrorAction Stop | format-table
}
catch
{
    Write-Host "$sharedInboxName is not a valid shared inbox! Check spelling." -ForegroundColor DarkRed
    exit
}

Write-Host "-----------------------------------------------------------------------------`n" -ForegroundColor DarkGray

Write-Host "Adding users to Shared Inbox...`n=======================================`n" -ForegroundColor Cyan

# Delete contents of file to ensure only output shown is from the current instance
if(Test-Path $PSScriptRoot\Incorrect-Names.txt -PathType Leaf)
{
    Clear-Content $PSScriptRoot\Incorrect-Names.txt
}

# Add users from CSV file to specified inbox
#=============================================================
foreach ($user in $users)
{
    $email = $user.'Email'
    $group = $sharedInboxName
    [bool] $hasFullAccess = $false
    [bool] $hasSendAs = $false

    # User email is valid
    if(Get-ADUser -Filter "EmailAddress -eq '$email'")
    {
        $selectedUser = Get-ADUser -Filter "EmailAddress -eq '$email'"

        # Check if user already has full access or send as permission
        if ((Get-MailboxPermission -Identity $sharedInboxName -User $email) -Or (Get-RecipientPermission -Identity $sharedInboxName -Trustee $email))
        {
            # User already has both full access and send as
            if ((Get-MailboxPermission -Identity $sharedInboxName -User $email) -And (Get-RecipientPermission -Identity $sharedInboxName -Trustee $email))
            {
                $hasFullAccess = $true
                $hasSendAs = $true

                Write-Host " WARN : $email already has Full Access & Send As permissions" -ForegroundColor Yellow

                continue
            }

            # User already has full access
            if (Get-MailboxPermission -Identity $sharedInboxName -User $email)
            {
                $hasFullAccess = $true

                Write-Host " WARN : $email already has Full Access permission" -ForegroundColor Yellow
            }
    
            # User already has send as
            if (Get-RecipientPermission -Identity $sharedInboxName -Trustee $email)
            {
                $hasSendAs = $true

                Write-Host " WARN : $email already has Send As permission" -ForegroundColor Yellow
            }
    
        }

        # Give user Full Access AND/OR Send As permissions depending on existing access
        try
        {
            # [void](<function>) hides the output of the command

            # Add full access if user is not already member
            if ($hasFullAccess -eq $false)
            {
                [void] ( Add-MailboxPermission "$group" -User $email -AccessRights FullAccess -InheritanceType all -ErrorAction Stop )	# Add user to Full Access
            }

            # Add send as if user is not already member
            if ($hasSendAs -eq $false)
            {
                [void] ( Add-RecipientPermission "$group" -Trustee $email -AccessRights SendAs -confirm:$false -ErrorAction Stop )		# Add user to Send As
            }
        }
        catch [System.Management.Automation.RemoteException]
        {
            Write-Host " WARN : Unexpected Exception" -ForegroundColor DarkRed
            continue
        }

        # Display what permissions were given to users
        # ----------------------------
        if (($hasFullAccess -eq $false) -And ($hasSendAs -eq $false))
        {
            Write-Host " ADDED: $email given Full Access & Send As for $group" -ForegroundColor Green

            continue
        }

        if ($hasFullAccess -eq $false)
        {
            Write-Host " ADDED: $email given Full Access for $group" -ForegroundColor Green
        }

        if ($hasSendAs -eq $false)
        {
            Write-Host " ADDED: $email given Send As for $group" -ForegroundColor Green
        }
        # ----------------------------

    }

    # User email is invalid
    else
    {
        Write-Host " ERROR: $email does not exist" -ForegroundColor DarkRed

        "$email" | Out-File -FilePath $PSScriptRoot\Incorrect-Names.txt -Append

        continue
    }
}
#=============================================================

Write-Host "`n=======================================`n" -ForegroundColor Cyan
Write-Host "Done!`n" -ForegroundColor Green

Write-Host "The following users were not found:" -ForegroundColor Red
Write-Host "=======================================`n" -ForegroundColor Red

if (Test-Path $PSScriptRoot\Incorrect-Names.txt -PathType Leaf)
{
    Get-Content $PSScriptRoot\Incorrect-Names.txt
}
else
{
    Write-Host "None!" -ForegroundColor Red
}

Write-Host "`n=======================================`n" -ForegroundColor Red

Write-Host "------------`nOutput saved to:`n$PSScriptRoot\Incorrect-Names.txt`n------------`n`n" -ForegroundColor Yellow


Write-Host "Disconnecting from MS Exchange..." -ForegroundColor Gray
Disconnect-ExchangeOnline -Confirm:$false