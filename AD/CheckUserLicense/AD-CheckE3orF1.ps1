#****************************************************
# Batch check if users in CSV have E3 or F1 license *
#****************************************************

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

# CSV
$users = Import-Csv -Delimiter "," -Path "$PSScriptRoot\Contacts.csv"

Write-Host "Checking users' licenses...`n=======================================`n" -ForegroundColor Cyan

# Delete contents of files to ensure only output shown is from the current instance
if (Test-Path $PSScriptRoot\Incorrect-Names.txt -PathType Leaf)
{
    Clear-Content $PSScriptRoot\Incorrect-Names.txt
}
if (Test-Path $PSScriptRoot\F1-User.txt -PathType Leaf)
{
    Clear-Content $PSScriptRoot\F1-User.txt
}
if (Test-Path $PSScriptRoot\E3-User.txt -PathType Leaf)
{
    Clear-Content $PSScriptRoot\E3-User.txt
}

# Check if users in CSV have E3 or F1 license
#=============================================================
foreach ($user in $users)
{
    $email = $user.'Email'
    $f1 = 'LIC_F1_Std'
    $e3 = 'LIC_E3_Std'

    # Check user email is valid
    #===========================
    if (-not (Get-ADUser -Filter "EmailAddress -eq '$email'")) # Email is not valid
    {
        Write-Host " ERROR: $email does not exist" -ForegroundColor DarkRed

        "$email" | Out-File -FilePath $PSScriptRoot\Incorrect-Names.txt -Append
        continue
    }

    # Extract SAM name, and check what LIC user has
    #==============================================
    $samName = Get-ADUser -Filter "EmailAddress -eq '$email'" -properties SamAccountName | Select-Object SamAccountName | % {$_.SamAccountName}
    $userGroups = Get-ADUser $samName -property * | select -expand memberof -Verbose
    
    if (($userGroups) -match $f1)
    {
        Write-Host " F1   : $email is an EmailOnly user" -ForegroundColor Yellow
        "$email" | Out-File -FilePath $PSScriptRoot\F1-User.txt -Append
    }
    elseif (($userGroups) -match $e3)
    {
        Write-Host " E3   : $email is a Full user" -ForegroundColor Green
        "$email" | Out-File -FilePath $PSScriptRoot\E3-User.txt -Append
    }

}
#=============================================================

Write-Host "`n=======================================`n" -ForegroundColor Cyan
Write-Host "Done!`n" -ForegroundColor Green


# Script output
#==============
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

#====

Write-Host "The following users have an E3 license:" -ForegroundColor Green
Write-Host "=======================================`n" -ForegroundColor Green

if (Test-Path $PSScriptRoot\E3-User.txt -PathType Leaf)
{
    Get-Content $PSScriptRoot\E3-User.txt
}
else
{
    Write-Host "None!" -ForegroundColor Green
}

Write-Host "`n=======================================`n" -ForegroundColor Green

#====

Write-Host "The following users have an F1 license:" -ForegroundColor Yellow
Write-Host "=======================================`n" -ForegroundColor Yellow

if (Test-Path $PSScriptRoot\F1-User.txt -PathType Leaf)
{
    Get-Content $PSScriptRoot\F1-User.txt
}
else
{
    Write-Host "None!" -ForegroundColor Yellow
}

Write-Host "`n=======================================`n" -ForegroundColor Yellow


# Output path
#============
Write-Host "------------`nOutput saved to:`n" -ForegroundColor Cyan

if (Test-Path $PSScriptRoot\Incorrect-Names.txt -PathType Leaf)
{
    Write-Host "$PSScriptRoot\Incorrect-Names.txt" -ForegroundColor Cyan
}
if (Test-Path $PSScriptRoot\E3-User.txt -PathType Leaf)
{
    Write-Host "$PSScriptRoot\E3-User.txt" -ForegroundColor Cyan
}
if (Test-Path $PSScriptRoot\F1-User.txt -PathType Leaf)
{
    Write-Host "$PSScriptRoot\F1-User.txt" -ForegroundColor Cyan    
}

Write-Host " `n------------`n`n" -ForegroundColor Cyan
