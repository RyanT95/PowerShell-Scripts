#*******************************************
# Batch add manual contacts to AD from CSV *
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
        Write-Host "Script needs to be run as admin!" -ForegroundColor Red
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

# CSV
$users = Import-Csv -Delimiter "," -Path "$PSScriptRoot\Contacts.csv"

# Delete contents of file to ensure only output shown is from the current instance
if(Test-Path $PSScriptRoot\Added-Contacts.txt -PathType Leaf)
{
    Clear-Content $PSScriptRoot\Added-Contacts.txt
}

Write-Host "Adding contacts...`n=======================================`n" -ForegroundColor Cyan

# Add users from CSV file as manual contacts in AD
foreach ($user in $users)
{
    $userFullname = $user.'Fullname'
    $userFirstname = $user.'First'
    $userSurname = $user.'Last'
    $email = $user.'Email'
    $desc = $user.'Desc'
    $ou = "OU=ManualContacts,OU=Contacts,dc=<DOMAIN-COMPONENT>,dc=co,dc=uk"

    New-ADObject -Name "$userFullname" -Type Contact -OtherAttributes @{DisplayName=$userFullname;givenName=$userFirstname;sn=$userSurname;mail=$email;Description=$desc} -Path $ou

    "$userFullname" | Out-File -FilePath $PSScriptRoot\Added-Contacts.txt -Append

    Write-Host " ADDED: $userFullname as contact" -ForegroundColor Green
}

Write-Host "`n=======================================`n" -ForegroundColor Cyan
Write-Host "Done!`n" -ForegroundColor Green

Write-Host "The following contacts were added:" -ForegroundColor Yellow
Write-Host "=======================================`n" -ForegroundColor Yellow

if (Test-Path $PSScriptRoot\Added-Contacts.txt -PathType Leaf)
{
    Get-Content $PSScriptRoot\Added-Contacts.txt
}
else
{
    Write-Host "None!" -ForegroundColor Yellow
}

Write-Host "`n=======================================`n" -ForegroundColor Yellow

Write-Host "Output saved to:`n$PSScriptRoot\Added-Contacts.txt`n`n" -ForegroundColor DarkYellow
