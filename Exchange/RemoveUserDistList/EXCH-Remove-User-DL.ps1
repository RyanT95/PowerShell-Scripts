#******************************************
# Remove user from all Distribution Lists *
#******************************************

# Check script is running as admin
#=============================================================
param([switch]$elevated)

# Get the current user, and check if they are an Administrator
Function Is-Admin 
{
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

# User is not admin
if ((Is-Admin) -eq $false)
{
    # If user doesn't have required permissions, present error and exit.
    # Otherwise, exit PowerShell and start new instance prompting user to log in as admin.
    if ($elevated) 
    {
        Write-Host "Script requires Admin permissions!" -ForegroundColor DarkRed
    }
    else
	{
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
	}
    
	exit
}

Write-Host "Running as admin...`n" -ForegroundColor Green
#=============================================================

#**************************
# Find inputted user in AD
#**************************
Function Get-User-Details
{
    # USER INPUT: user email
    Write-Host "`nInput full email address of user : " -ForegroundColor Green -NoNewline
    $email = Read-Host

    Write-Host "`n-----------------------------------------------------------------------------`n" -ForegroundColor DarkGray -NoNewline
    Write-Host "USER DETAILS`n" -ForegroundColor DarkGray -NoNewline
    Write-Host "-----------------------------------------------------------------------------`n" -ForegroundColor DarkGray -NoNewline

    #Can be shared inbox, but can't be DL#

    # Check user email is valid
    #===========================
    if(-not (Get-ADUser -Filter "EmailAddress -eq '$email'" -ErrorAction SilentlyContinue)) # Email is NOT user
    {
        if(-not (Get-DistributionGroup -Identity $email -ErrorAction SilentlyContinue)) # Email is NOT DL
        {
            if(-not (Get-Mailbox -RecipientTypeDetails SharedMailbox -Identity $email -ErrorAction SilentlyContinue)) # Email is NOT mailbox
            {
                Write-Host " ERROR: $email does not exist" -ForegroundColor DarkRed
                
                Disconnect-ExchangeOnline -Confirm:$false
                continue
            }
            else # Email IS mailbox
            {
                $isMailbox = $true
                $selectedUser = Get-Mailbox -RecipientTypeDetails SharedMailbox -Identity $email
            }
        }
        else # Email IS DL
        {
            Write-Host " ERROR: email cannot be Distribution List" -ForegroundColor DarkRed

            Disconnect-ExchangeOnline -Confirm:$false
            continue
        }   
    }
    else # Email IS user
    {
        $isMailbox = $false
        $selectedUser = Get-ADUser -Filter "EmailAddress -eq '$email'"
    }

    if ($isMailbox -eq $false)
    {
        (Get-ADUser -Filter "EmailAddress -eq '$email'" -Properties DisplayName, Title, EmailAddress | Format-List DisplayName, Title, EmailAddress| Out-String).trim()
    }

    if ($isMailbox -eq $true)
    {
        (Get-Mailbox -RecipientTypeDetails SharedMailbox -Identity $email | Format-List DisplayName, RecipientTypeDetails, PrimarySmtpAddress | Out-String).trim()
    }
    
    #===========================

    Remove-User-DL($email)
}

#**************************
# Remove user from all DLs
#**************************
Function Remove-User-DL($Email)
{
    Write-Host "`n-----------------------------------------------------------------------------`n" -ForegroundColor DarkGray -NoNewline
    Write-Host "USER DISTRIBUTION GROUP MEMBERSHIPS`n" -ForegroundColor DarkGray -NoNewline
    Write-Host "-----------------------------------------------------------------------------`n" -ForegroundColor DarkGray -NoNewline

    $dn = Get-User $Email | select -ExpandProperty DistinguishedName
    $groups = Get-Group -Filter "Members -eq '$dn'" -RecipientTypeDetails MailUniversalDistributionGroup | Select-Object -ExpandProperty DisplayName #| Out-String

    foreach ($group in $groups)
    {
        Write-Host $group
    }

    # Confirm removal of DLs
    #===========================
    do
    {
        Write-Host "`nConfirm Removal of Groups? (y/n) : " -ForegroundColor Green -NoNewline

        $confirm = Read-Host

        switch ($confirm)
	    {
	    	y {Write-Host "Proceeding.`n"   -ForegroundColor Gray;  $valid = $true;  continue}
            Y {Write-Host "Proceeding.`n"   -ForegroundColor Gray;  $valid = $true;  continue}
	    	
            n {Write-Host "Cancelling...`n" -ForegroundColor Gray;  $valid = $true;  return}
            N {Write-Host "Cancelling...`n" -ForegroundColor Gray;  $valid = $true;  return}
            
	    	default {Write-Host "Invalid Selection! Try again.`n" -ForegroundColor DarkRed; $valid = $false;}
	    }
    }
    until ($valid -eq $true)
    #===========================

    Write-Host "`nRemoving user from Distribution Lists...`n=======================================`n" -ForegroundColor Cyan

    # Remove user from all Distribution Groups
    foreach ($group in $groups)
    {
        #Remove-DistributionGroupMember -Identity "$group" -Member "$Email" -Confirm:$false

        Write-Host "Removed : $Email from $group"
    }

    Write-Host "`n=======================================`n" -ForegroundColor Cyan

    Write-Host "Done!`n" -ForegroundColor Green
}

#*************
# Entry Point
#*************

# Get user who ran script
$admCred = $env:UserName + "@DOMAIN.com"

# Connect to MS Exchange
Import-Module ExchangeOnlineManagement -Verbose:$false
Connect-ExchangeOnline -UserPrincipalName $admCred -ShowBanner:$false -ShowProgress:$true -Verbose:$false

Get-User-Details

Write-Host "Disconnecting from MS Exchange..." -ForegroundColor Gray
Disconnect-ExchangeOnline -Confirm:$false
