#************************************************
# Batch add users to distribution list from CSV *
#************************************************

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
# Add users from CSV to DL
#**************************
Function Rem-Users
{
    # USER INPUT: distribution list name
    $distGroupName = Read-Host -Prompt 'Input full email address of distribution list'
    
    Write-Host "`n-----------------------------------------------------------------------------`n" -ForegroundColor DarkGray
    
    # Error checking for valid distribution list
    try
    {
        Get-DistributionGroup -Identity $distGroupName -ErrorAction Stop | format-table
    }
    catch
    {
        Write-Host "$distGroupName is not a valid distribution list! Check spelling." -ForegroundColor DarkRed
        exit
    }
    
    Write-Host "-----------------------------------------------------------------------------`n" -ForegroundColor DarkGray
    
    Write-Host "Removing users from Distribution List...`n=======================================`n" -ForegroundColor Cyan
    
    # Delete contents of file to ensure only output shown is from the current instance
    if(Test-Path $PSScriptRoot\Incorrect-Names.txt -PathType Leaf)
    {
        Clear-Content $PSScriptRoot\Incorrect-Names.txt
    }
    
    # Add users from CSV file to specified DL
    #=============================================================
    foreach ($user in $users)
    {
        $email = $user.'Email'
        $group = $distGroupName
    
        # Check user email is valid
        #===========================
        if(-not ( Get-ADUser -Filter "EmailAddress -eq '$email'" -ErrorAction SilentlyContinue)) # Email is NOT user
        {
            if(-not (Get-DistributionGroup -Identity $email -ErrorAction SilentlyContinue)) # Email is NOT DL
            {
                if(-not (Get-Mailbox -RecipientTypeDetails SharedMailbox -Identity $email -ErrorAction SilentlyContinue)) # Email is NOT mailbox
                {
					if(-not (Get-MailContact -Identity $email -ErrorAction SilentlyContinue)) # Email is NOT contact
                    {
						Write-Host " ERROR   : $email does not exist" -ForegroundColor DarkRed
		
						"$email" | Out-File -FilePath $PSScriptRoot\Incorrect-Names.txt -Append
		
						continue
					}
					else # Email IS contact
					{
						$selectedUser = Get-MailContact -Identity $email
					}
                }
                else # Email IS mailbox
                {
                    $selectedUser = Get-Mailbox -RecipientTypeDetails SharedMailbox -Identity $email
                }
            }
            else # Email IS DL
            {
                $selectedUser = Get-DistributionGroup -Identity $email
            }   
        }
        else # Email IS user
        {
            $selectedUser = Get-ADUser -Filter "EmailAddress -eq '$email'"
        }
		
    
        # Check if user is a member of the DL
        #=====================================
        $groupMembers = Get-DistributionGroupMember -Identity $distGroupName
        
        if(($selectedUser) -ilike (Get-DistributionGroup -Identity $distGroupName))
        {
            Write-Host " WARN    : Cannot add $email to itself!" -ForegroundColor Red
            continue         
        }

        if(-not ("$groupMembers".Contains($selectedUser.Name)))
        {
            Write-Host " WARN    : $email is not a member of the DL" -ForegroundColor Yellow
            continue
        }
    
    
        # Remove user from DL
        #====================
        try
        {
            Remove-DistributionGroupMember -Identity "$group" -member $email -BypassSecurityGroupManagerCheck -ErrorAction Stop -Confirm:$false
        }
        catch [System.Management.Automation.RemoteException]
        {
            Write-Host " WARN    : Unexpected Exception" -ForegroundColor DarkRed -NoNewLine
			Write-Host " when removing: $email" -ForegroundColor Gray
			
            continue
        }
    
        Write-Host " REMOVED : $email from $group" -ForegroundColor Green
    }
    #=============================================================
    
    Write-Host "`n=======================================`n" -ForegroundColor Cyan
    Write-Host "Done!`n" -ForegroundColor Green
}

#************************
# Show results of script
#************************
Function Show-Results
{
    Write-Host "The following users were not found:" -ForegroundColor Red
    Write-Host "=======================================`n" -ForegroundColor Red
    
    Get-Content $PSScriptRoot\Incorrect-Names.txt
    
    Write-Host "`n=======================================`n" -ForegroundColor Red
    
    Write-Host "------------`nOutput saved to:`n$PSScriptRoot\Incorrect-Names.txt`n------------`n`n" -ForegroundColor Cyan
}


#*************
# Entry point
#*************

# USER INPUT
$domain = Read-Host -Prompt 'Input company domain name (i.e. @domain.com)'

# Get user who ran script
$admCred = $env:UserName + $domain

# Connect to MS Exchange
Import-Module ExchangeOnlineManagement -Verbose:$false
Connect-ExchangeOnline -UserPrincipalName $admCred -ShowBanner:$false -ShowProgress:$true -Verbose:$false

# CSV
$users = Import-Csv -Delimiter "," -Path "$PSScriptRoot\Contacts.csv"

Rem-Users
Show-Results

Write-Host "Disconnecting from MS Exchange..." -ForegroundColor Gray
Disconnect-ExchangeOnline -Confirm:$false