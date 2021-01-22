#************************************************
# Batch add users to distribution list from CSV *
#************************************************

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

Write-Host "Adding users to Distribution List...`n=======================================`n" -ForegroundColor Cyan

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
                Write-Host " ERROR: $email does not exist" -ForegroundColor DarkRed

                "$email" | Out-File -FilePath $PSScriptRoot\Incorrect-Names.txt -Append

                continue
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


    # Check if user is already a member of the DL
    #============================================
    $groupMembers = Get-DistributionGroupMember -Identity $distGroupName
    
    if("$groupMembers".Contains($selectedUser.Name))
    {
        if(($selectedUser) -ilike (Get-DistributionGroup -Identity $distGroupName))
        {
            Write-Host " WARN : Cannot add $email to itself!" -ForegroundColor Red
            continue         
        }

        Write-Host " WARN : $email is already a member" -ForegroundColor Yellow
        continue
    }


    # Add user to DL
    #===============
    try
    {
        Add-DistributionGroupMember -Identity "$group" -member $email -BypassSecurityGroupManagerCheck -ErrorAction Stop
    }
    catch [System.Management.Automation.RemoteException]
    {
        Write-Host " WARN : Unexpected Exception" -ForegroundColor DarkRed
        continue
    }

    Write-Host " ADDED: $email to $group" -ForegroundColor Green
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

Write-Host "------------`nOutput saved to:`n$PSScriptRoot\Incorrect-Names.txt`n------------`n`n" -ForegroundColor Cyan


Write-Host "Disconnecting from MS Exchange..." -ForegroundColor Gray
Disconnect-ExchangeOnline -Confirm:$false